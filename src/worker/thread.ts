/**
 * A thread of the pool (`docs/async.md`): it holds sessions as the same thread does and runs one task at a time in
 * portions, so a `cancel` gets through between them, and inside a formula through memory shared with the pool. A
 * session's changing calls carry sequence numbers: the last one again gets the same answer, so a repeated message
 * never applies an edit twice. A restore opens a snapshot and replays the journal with the clocks the session had.
 */
import { callLimits } from '../async/options.js'
import { budgeted, openInlineSession, type InlineDefaults, type InlineSession } from '../async/session.js'
import { computeSteps } from '../compute.js'
import type { AbortSignalLike, BookInput, CalcOptions, Clock, ExternalAnswer, ExternalResolver, StructureOp } from '../contract.js'
import { currentClock } from '../env.js'
import { abortable, CalcError } from '../exec/budget.js'
import { checkComputeOptions } from '../input.js'
import { registerLocale } from '../locale/load.js'
import { LOCALES } from '../locale/registry.js'
import type { Locale } from '../locale/types.js'
import { heapCaches } from '../range-cache.js'
import type { HistoryData } from '../session/undo.js'
import { GEOMETRY } from '../structure/geometry.js'
import { CHANGING, EDITS, toData, type Answer, type Entry, type Op, type ToThread } from './protocol.js'
import { LOCALE_MODULES } from '../locale/modules.js'
import { setRegexLevel } from '../regex-level.js'
import type { FunctionGroupName } from '../contract.js'
import { GROUP_ARITY } from '../groups/manifest.js'
import { GROUP_MODULES } from '../groups/modules.js'
import { loadGroups, useManifest } from '../groups/registry.js'
import type { RegexLevel } from '../contract.js'
import { workerThreads } from './platform.js'
import { pack, packData, unpack, unpackCached, unpackData, type Sent } from './pack.js'

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
const waiting: Exclude<ToThread, { t: 'init' | 'locales' | 'regex' | 'functions' | 'cancel' | 'answer' }>[] = []
let current: { task: number; abort(reason?: unknown): void } | null = null
/** Where the pool writes the task it cancels (`init`), none where threads share no memory: the engine's checks read it inside a formula. */
let cancelled: Int32Array | undefined
const CANCELLED = 'cancelled by the pool'

/** The signal of task `task`: aborted by the pool's `cancel`, or as soon as a check reads the task in the shared word. */
function taskSignal(task: number): { signal: AbortSignalLike; abort(reason?: unknown): void } {
  const ac = abortable()
  const word = cancelled
  if (!word) return ac
  const s = ac.signal
  const signal: AbortSignalLike = {
    get aborted() {
      if (!s.aborted && Atomics.load(word, 0) === task) ac.abort(CANCELLED)
      return s.aborted
    },
    get reason() {
      return s.reason
    },
    addEventListener: s.addEventListener,
    removeEventListener: s.removeEventListener,
  }
  return { signal, abort: ac.abort }
}
/** The locales and the engine the pool sent, loaded before the next task. */
let loading: Promise<void> = Promise.resolve()

/** Loads locales the pool sent: a package's from its own module, with its messages; another, or one whose module did not load, as sent. */
function load(list: Locale[]) {
  loading = loading
    .then(async () => {
      for (const sent of list) if (!LOCALES.has(sent.tag)) registerLocale((await LOCALE_MODULES[sent.tag]?.().catch(() => undefined)) ?? sent)
    })
    .catch(() => {})
}

/** Loads the level of match mode 3 the pool has registered, its module imported by the thread (a functions' object no message carries), before the next task. */
function loadRegex(kind: RegexLevel['kind']) {
  loading = loading
    .then(async () => {
      const m = await import('../regex/index.js')
      setRegexLevel(kind === 'xlcalc-pcre-syntax' ? m.pcreSyntax : m.pcreEngine)
    })
    .catch(() => {})
}

// a thread starts without the groups of functions and loads those the pool names: the root's all, the light entry's registered
useManifest(GROUP_ARITY)

/** Loads the groups of functions the pool names, each from its own module, before the next task. */
function loadFunctions(names: FunctionGroupName[]) {
  loading = loading
    .then(async () => loadGroups(await Promise.all(names.map((n) => GROUP_MODULES[n]()))))
    .catch(() => {})
}

/** A session call by name; `o` goes where the call takes options, with the edit options a changing call carries last (`EDITS`). */
async function invoke(h: Held, op: Op, args: unknown[], o: CalcOptions): Promise<Answer> {
  const s = h.s
  const n = EDITS[op]
  const eo = n === undefined ? o : { ...o, ...(args[n] as object | undefined) }
  let value: unknown
  switch (op) {
    case 'editStructure':
    case 'undo':
    case 'redo': {
      // a structure operation's mapCell goes as its geometry, of which the calling side makes the same function
      const r = op === 'editStructure' ? await s.editStructure(args[0] as StructureOp, eo) : await s[op](eo)
      const { mapCell, ...rest } = (r ?? {}) as { mapCell?: object }
      value = r && { ...rest, ...(mapCell ? { geometry: GEOMETRY.get(mapCell) } : {}) }
      break
    }
    case 'getBook':
      value = await pack(await s.getBook())
      break
    case 'getLayout':
      value = await s.getLayout()
      break
    case 'getUsedRange':
      value = await s.getUsedRange(args[0] as string)
      break
    case 'getAutoFilterRange':
      value = await s.getAutoFilterRange(args[0] as string)
      break
    case 'getEnd':
      value = await s.getEnd(args[0] as string, args[1] as string, args[2] as never)
      break
    case 'getAutoSum':
      value = await s.getAutoSum(args[0] as string, args[1] as string)
      break
    case 'getNamesFromSelection':
      value = await s.getNamesFromSelection(args[0] as string, args[1] as string, args[2] as never)
      break
    case 'getSpecialCells':
      value = await s.getSpecialCells(args[0] as string, args[1] as string, args[2] as never)
      break
    case 'getWorkbookStatistics':
      value = await s.getWorkbookStatistics(args[0] as string)
      break
    case 'explain':
      value = await s.explain(args[0] as string, args[1] as string, { ...(args[2] as object | undefined), ...o })
      break
    case 'getResult':
      value = await s.getResult(o)
      break
    case 'getSnapshot': {
      const r = await s.getSnapshot(o)
      value = { book: await pack(r.book), result: r.result }
      break
    }
    case 'getHash':
      value = await s.getHash(o)
      break
    case 'getCells':
      value = await s.getCells(args[0] as string, args[1] as string | string[], o)
      break
    case 'copyText':
      value = await s.copyText(args[0] as string, args[1] as string, { ...(args[2] as object | undefined), ...o })
      break
    case 'getValues':
      value = await s.getValues(args[0] as string, args[1] as string, { ...(args[2] as object | undefined), ...o })
      break
    case 'getCell':
      value = await s.getCell(args[0] as string, args[1] as string)
      break
    case 'getViewport':
      value = await s.getViewport(args[0] as string, args[1] as string, args[2] as never)
      break
    case 'getCellIssue':
      value = await s.getCellIssue(args[0] as string, args[1] as string)
      break
    case 'getValidation':
      value = await s.getValidation(args[0] as string, args[1] as string)
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
  h.s = await openInlineSession(book, opts, d, { clock: () => h.planned ?? currentClock(), history, revision, copySnapshot: false })
  held.set(sid, h)
  return h
}

/** A session as a snapshot takes it: its book and history as text, its clock, its revision and its last changing call. */
const taken = async (h: Held) => ({ book: await pack(await h.s.getBook()), clock: h.s.clock, seq: h.seq, history: packData(h.s.historyData()), revision: h.s.revision })

/** Options with the stored values a message carries as text (`ComputeOptions.cached`), as objects again. */
const withCached = <T extends object>(opts: T, cached: Sent | undefined): T => (cached === undefined ? opts : { ...opts, cached: unpackCached(cached) })

/** The calls outside the workbook sent to the pool's resolver, by number, till their answers come. */
const asked = new Map<number, (m: { value?: ExternalAnswer; failed?: true }) => void>()
let asks = 0

/** A resolver asking the pool's for session `sid`, or for the computation of task `task`; the call's abort reaches it as `unask`. */
const asking =
  (to: { sid: number } | { task: number }): ExternalResolver =>
  ({ signal, ...call }) => {
    if (signal.aborted) return Promise.reject(signal.reason)
    const ask = ++asks
    return new Promise<ExternalAnswer>((resolve, reject) => {
      const unask = () => {
        asked.delete(ask)
        port.postMessage({ unask: ask })
        reject(signal.reason)
      }
      asked.set(ask, (m) => {
        signal.removeEventListener('abort', unask)
        asked.delete(ask)
        if (m.failed) reject(new Error('the resolver gave no answer'))
        else resolve(m.value)
      })
      signal.addEventListener('abort', unask, { once: true })
      port.postMessage({ ask, ...to, call })
    })
  }

/** Options whose `externals` answer by the pool's resolver where the message says it has one (`asks`). */
const withAsks = <T extends object>(opts: T, asks: true | undefined, to: { sid: number } | { task: number }): T =>
  asks ? { ...opts, externals: { ...(opts as { externals?: object }).externals, resolve: asking(to) } } : opts

async function run(m: (typeof waiting)[number], signal: AbortSignalLike): Promise<unknown> {
  switch (m.t) {
    case 'compute': {
      const c = checkComputeOptions(withAsks(withCached(m.opts, m.cached), m.asks, { task: m.task }))
      return budgeted(d, callLimits(d.limits, c.limits))(computeSteps(unpack(m.book), c), { signal, timeoutMs: m.timeoutMs })
    }
    case 'open': {
      const h = await open(m.sid, unpack(m.book), withAsks(withCached({ ...m.opts, signal, timeoutMs: m.timeoutMs }, m.cached), m.asks, { sid: m.sid }), null)
      return { value: null, clock: h.s.clock, history: h.s.history, revision: h.s.revision } satisfies Answer
    }
    case 'call': {
      const h = heldOf(m.sid)
      if (m.seq !== undefined) {
        if (m.seq === h.seq && h.last) return h.last
        if (m.seq !== h.seq + 1) throw new CalcError('thread_failed', `call ${m.seq} of session ${m.sid} out of order (the last was ${h.seq})`)
      }
      // a read without options is bounded by the thread's default, as createInlineCalc's reads are by its own: here the
      // call's deadline (the pool's timeoutMs), for this call alone, as the thread runs one task at a time
      d.timeoutMs = m.timeoutMs
      let a: Answer
      try {
        a = await invoke(h, m.op, m.args, { signal, timeoutMs: m.timeoutMs })
      } finally {
        d.timeoutMs = undefined
      }
      if (m.seq !== undefined && (CHANGING as readonly string[]).includes(m.op)) {
        h.seq = m.seq
        h.last = a
      }
      return a
    }
    case 'restore': {
      const { snapshot, journal } = m
      held.get(m.sid)?.s.kill(new CalcError('closed', 'the session is restored'))
      const h = await open(m.sid, unpack(snapshot.book), withAsks(withCached({ ...snapshot.opts, signal }, snapshot.cached), m.asks, { sid: m.sid }), snapshot.clock, snapshot.history && unpackData(snapshot.history), snapshot.revision)
      h.seq = snapshot.seq
      try {
        for (const e of journal as Entry[]) {
          h.planned = e.clock
          const a = await invoke(h, e.op, e.args, { signal })
          if (a.clock.now !== e.clock.now || a.clock.seed !== e.clock.seed) throw new CalcError('restore_failed', `the replay of call ${e.seq} used a different clock`)
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
  const ac = taskSignal(m.task)
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
    if (m.regex) loadRegex(m.regex)
    loadFunctions(m.functions)
    cancelled = m.cancelled
  } else if (m.t === 'locales') {
    load(m.locales)
  } else if (m.t === 'regex') {
    loadRegex(m.kind)
  } else if (m.t === 'functions') {
    loadFunctions(m.functions)
  } else if (m.t === 'cancel') {
    if (current?.task === m.task) current.abort(CANCELLED)
  } else if (m.t === 'answer') {
    asked.get(m.ask)?.(m)
  } else {
    waiting.push(m)
    if (!current) void next()
  }
}

if (port.on) port.on('message', receive)
else port.addEventListener!('message', (e) => receive(e.data))
