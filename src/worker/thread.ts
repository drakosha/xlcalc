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
import { isHiddenRow } from '../layout/hidden.js'
import { heapCaches } from '../range-cache.js'
import { CHANGING, toData, type Answer, type Entry, type Op, type Outline, type ToThread } from './protocol.js'
import { workerThreads } from './platform.js'
import { pack, unpack } from './pack.js'

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
const waiting: Exclude<ToThread, { t: 'init' | 'cancel' }>[] = []
let current: { task: number; abort(reason?: unknown): void } | null = null

/** What `applyOp` reads of a book for `mapCell`: the sheets' names and visibility, their hidden rows and filter mode, their tables (a table deleted takes its cells). */
function outlineOf(book: BookInput): Outline {
  return {
    sheets: book.sheets.map((s) => {
      const rows = Object.entries(s.rows ?? {}).filter(([, r]) => isHiddenRow(r))
      return {
        name: s.name,
        cells: {},
        ...(rows.length ? { rows: Object.fromEntries(rows.map(([k]) => [k, { hidden: true }])) } : {}),
        ...(s.filtered ? { filtered: true } : {}),
        ...(s.tables ? { tables: s.tables } : {}),
        ...(s.hidden ? { hidden: s.hidden } : {}),
      }
    }),
    ...(book.names ? { names: book.names } : {}),
  }
}

/** A session call by name; `o` goes where the call takes options. */
async function invoke(h: Held, op: Op, args: unknown[], o: CalcOptions): Promise<Answer> {
  const s = h.s
  let value: unknown
  let outline: Outline | undefined
  switch (op) {
    case 'restructure': {
      outline = outlineOf(await s.book())
      const r = await s.restructure(args[0] as StructureOp, o)
      value = r.ok ? { ok: true, changes: r.changes } : r
      break
    }
    case 'setStyles':
      value = await s.setStyles(args[0] as never)
      break
    case 'book':
      value = await pack(await s.book())
      break
    case 'explain':
      value = await s.explain(args[0] as string, args[1] as string, args[2] as never)
      break
    case 'result':
      value = await s.result(o)
      break
    case 'hash':
      value = await s.hash(o)
      break
    case 'cells':
      value = await s.cells(args[0] as string, args[1] as string)
      break
    case 'validation':
      value = await s.validation(args[0] as string, args[1] as string)
      break
    case 'validateEntry':
      value = await s.validateEntry(args[0] as string, args[1] as string, args[2] as never, o)
      break
    default:
      value = await (s[op] as (a: unknown, o: CalcOptions) => Promise<unknown>)(args[0], o)
  }
  return { value, clock: s.clock, ...(outline ? { outline } : {}) }
}

function heldOf(sid: number): Held {
  const h = held.get(sid)
  if (!h) throw new CalcError('thread_failed', `session ${sid} is not on this thread`)
  return h
}

async function open(sid: number, book: BookInput, opts: object, planned: Clock | null): Promise<Held> {
  const h = { planned, seq: 0 } as Held
  h.s = await openInlineSession(book, opts, d, { clock: () => h.planned ?? currentClock() })
  held.set(sid, h)
  return h
}

async function run(m: (typeof waiting)[number], signal: AbortSignalLike): Promise<unknown> {
  switch (m.t) {
    case 'compute': {
      const c = checkComputeOptions(m.opts)
      return budgeted(d, callLimits(d.limits, c.limits))(computeSteps(unpack(m.book), c), { signal, timeoutMs: m.timeoutMs })
    }
    case 'open': {
      const h = await open(m.sid, unpack(m.book), { ...m.opts, signal, timeoutMs: m.timeoutMs }, null)
      return { value: null, clock: h.s.clock } satisfies Answer
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
      const h = await open(m.sid, unpack(snapshot.book), { ...snapshot.opts, signal }, snapshot.clock)
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
      return { book: await pack(await h.s.book()), clock: h.s.clock, seq: h.seq }
    }
    case 'drop': {
      const h = held.get(m.sid)
      const kept = m.keep && h ? { book: await pack(await h.s.book()), clock: h.s.clock, seq: h.seq } : null
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
  } else if (m.t === 'cancel') {
    if (current?.task === m.task) current.abort('cancelled by the pool')
  } else {
    waiting.push(m)
    if (!current) void next()
  }
}

if (port.on) port.on('message', receive)
else port.addEventListener!('message', (e) => receive(e.data))
