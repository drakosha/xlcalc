/**
 * A thread of the pool (`docs/async.md`): it holds sessions as the same thread does and runs one task at a time in
 * portions, so a `cancel` gets through between them. A session's changing calls carry sequence numbers: the last
 * one again gets the same answer, so a repeated message never applies an edit twice. A restore opens a snapshot and
 * replays the journal with the clocks the session had.
 */
import { abortable } from '../async/queue.js'
import { callLimits } from '../async/options.js'
import { budgeted, openInlineSession, type InlineDefaults, type InlineSession } from '../async/session.js'
import { computeSteps } from '../compute.js'
import type { AbortSignalLike, BookInput, CalcOptions, Clock, StructureOp } from '../contract.js'
import { currentClock } from '../env.js'
import { CalcError } from '../exec/budget.js'
import { checkComputeOptions } from '../input.js'
import { registerLocale } from '../locale/load.js'
import { LOCALES } from '../locale/registry.js'
import type { Locale } from '../locale/types.js'
import { heapCaches } from '../range-cache.js'
import type { HistoryData } from '../session/undo.js'
import { GEOMETRY } from '../structure/geometry.js'
import { CHANGING, EDITS, toData, type Answer, type Entry, type Op, type ToThread } from './protocol.js'
import { LOCALE_MODULES } from './locales.js'
import { workerThreads } from './platform.js'
import { pack, packData, unpack, unpackData } from './pack.js'

interface Port {
  postMessage(m: unknown): void
  on?(event: 'message', f: (m: ToThread) => void): void
  addEventListener?(event: 'message', f: (e: { data: ToThread }) => void): void
}

interface Held {
  s: InlineSession
  /** The last changing call applied, and its answer. */
  seq: number
  last?: Answer
  /** The clock the session takes next, while a restore replays; the current moment otherwise. */
  planned: Clock | null
}

const port = ((workerThreads() as { parentPort?: Port } | undefined)?.parentPort ?? globalThis) as Port
let d: InlineDefaults = { sliceMs: 50 }
const held = new Map<number, Held>()
const waiting: Exclude<ToThread, { t: 'init' | 'locales' | 'cancel' }>[] = []
let current: { task: number; abort(reason?: unknown): void } | null = null
/** The locales the pool sent, loaded before the next task. */
let loading: Promise<void> = Promise.resolve()

/** Loads locales the pool sent: a package's from its own module, with its messages; another, or one whose module did not load, as sent. */
function load(list: Locale[]) {
  loading = loading
    .then(async () => {
      for (const sent of list) if (!LOCALES.has(sent.tag)) registerLocale((await LOCALE_MODULES[sent.tag]?.().catch(() => undefined)) ?? sent)
    })
    .catch(() => {})
}

/** A session call by name; `o` goes where the call takes options, with the edit options a changing call carries last (`EDITS`). */
async function invoke(h: Held, op: Op, args: unknown[], o: CalcOptions): Promise<Answer> {
  const s = h.s
  const n = EDITS[op]
  const eo = n === undefined ? o : { ...o, ...(args[n] as object | undefined) }
  let value: unknown
  switch (op) {
    case 'restructure':
    case 'undo':
    case 'redo': {
      // a structure operation's mapCell goes as its geometry, of which the calling side makes the same function
      const r = op === 'restructure' ? await s.restructure(args[0] as StructureOp, eo) : await s[op](eo)
      const { mapCell, ...rest } = (r ?? {}) as { mapCell?: object }
      value = r && { ...rest, ...(mapCell ? { geometry: GEOMETRY.get(mapCell) } : {}) }
      break
    }
    case 'book':
      value = await pack(await s.book())
      break
    case 'layout':
      value = await s.layout()
      break
    case 'extent':
      value = await s.extent(args[0] as string)
      break
    case 'explain':
      value = await s.explain(args[0] as string, args[1] as string, args[2] as never)
      break
    case 'result':
      value = await s.result(o)
      break
    case 'getSnapshot': {
      const r = await s.getSnapshot(o)
      value = { book: await pack(r.book), result: r.result }
      break
    }
    case 'hash':
      value = await s.hash(o)
      break
    case 'cells':
      value = await s.cells(args[0] as string, args[1] as string | string[])
      break
    case 'getViewport':
      value = await s.getViewport(args[0] as string, args[1] as string, args[2] as never)
      break
    case 'getCellIssue':
      value = await s.getCellIssue(args[0] as string, args[1] as string)
      break
    case 'validation':
      value = await s.validation(args[0] as string, args[1] as string)
      break
    case 'validateEntry':
      value = await s.validateEntry(args[0] as string, args[1] as string, args[2] as never, o)
      break
    default:
      value = await (s[op] as (...a: unknown[]) => Promise<unknown>)(...args.slice(0, n), eo)
  }
  return { value, clock: s.clock, history: s.history, revision: s.revision }
}

function heldOf(sid: number): Held {
  const h = held.get(sid)
  if (!h) throw new CalcError('thread_failed', `session ${sid} is not on this thread`)
  return h
}

async function open(sid: number, book: BookInput, opts: object, planned: Clock | null, history?: HistoryData, revision?: number): Promise<Held> {
  const h = { planned, seq: 0 } as Held
  h.s = await openInlineSession(book, opts, d, { clock: () => h.planned ?? currentClock(), history, revision })
  held.set(sid, h)
  return h
}

/** A session as a snapshot takes it: its book and history as text, its clock, its revision and its last changing call. */
const taken = async (h: Held) => ({ book: await pack(await h.s.book()), clock: h.s.clock, seq: h.seq, history: packData(h.s.historyData()), revision: h.s.revision })

async function run(m: (typeof waiting)[number], signal: AbortSignalLike): Promise<unknown> {
  switch (m.t) {
    case 'compute': {
      const c = checkComputeOptions(m.opts)
      return budgeted(d, callLimits(d.limits, c.limits))(computeSteps(unpack(m.book), c), { signal, timeoutMs: m.timeoutMs })
    }
    case 'open': {
      const h = await open(m.sid, unpack(m.book), { ...m.opts, signal, timeoutMs: m.timeoutMs }, null)
      return { value: null, clock: h.s.clock, history: h.s.history, revision: h.s.revision } satisfies Answer
    }
    case 'call': {
      const h = heldOf(m.sid)
      if (m.seq !== undefined) {
        if (m.seq === h.seq && h.last) return h.last
        if (m.seq !== h.seq + 1) throw new CalcError('thread_failed', `call ${m.seq} of session ${m.sid} out of order (the last was ${h.seq})`)
      }
      const a = await invoke(h, m.op, m.args, { signal, timeoutMs: m.timeoutMs })
      if (m.seq !== undefined && (CHANGING as readonly string[]).includes(m.op)) {
        h.seq = m.seq
        h.last = a
      }
      return a
    }
    case 'restore': {
      const { snapshot, journal } = m
      held.get(m.sid)?.s.kill(new CalcError('closed', 'the session is restored'))
      const h = await open(m.sid, unpack(snapshot.book), { ...snapshot.opts, signal }, snapshot.clock, snapshot.history && unpackData(snapshot.history), snapshot.revision)
      h.seq = snapshot.seq
      try {
        for (const e of journal as Entry[]) {
          h.planned = e.clock
          const a = await invoke(h, e.op, e.args, { signal })
          if (a.clock.now !== e.clock.now || a.clock.seed !== e.clock.seed) throw new CalcError('restore_failed', `the replay of call ${e.seq} took another clock`)
          h.seq = e.seq
          h.last = a
        }
      } catch (e) {
        // a restore that fails leaves nothing of the session behind: the pool closes it
        h.s.kill(new CalcError('closed', 'the session could not be restored'))
        held.delete(m.sid)
        throw e
      }
      h.planned = null
      return null
    }
    case 'snapshot': {
      const h = heldOf(m.sid)
      return taken(h)
    }
    case 'drop': {
      const h = held.get(m.sid)
      const kept = m.keep && h ? await taken(h) : null
      h?.s.kill(new CalcError('closed', 'the session left the thread'))
      held.delete(m.sid)
      return kept
    }
  }
}

async function next() {
  const m = waiting.shift()
  if (!m) return
  const ac = abortable()
  current = { task: m.task, abort: ac.abort }
  try {
    await loading
    const value = await run(m, ac.signal)
    try {
      port.postMessage({ task: m.task, ok: true, value })
    } catch (e) {
      port.postMessage({ task: m.task, ok: false, error: toData(e) })
    }
  } catch (e) {
    port.postMessage({ task: m.task, ok: false, error: toData(e) })
  } finally {
    current = null
    void next()
  }
}

function receive(m: ToThread) {
  if (m.t === 'init') {
    d = { sliceMs: m.sliceMs, limits: m.limits }
    // the thread's sessions and computations share its heap, and so one budget of their caches
    if (m.cacheBytes) heapCaches.cap = m.cacheBytes
    if (m.locales) load(m.locales)
  } else if (m.t === 'locales') {
    load(m.locales)
  } else if (m.t === 'cancel') {
    if (current?.task === m.task) current.abort('cancelled by the pool')
  } else {
    waiting.push(m)
    if (!current) void next()
  }
}

if (port.on) port.on('message', receive)
else port.addEventListener!('message', (e) => receive(e.data))
