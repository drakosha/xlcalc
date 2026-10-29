/**
 * The pool of threads (`docs/async.md`). A session is placed on a thread and stays there; one-off computations wait
 * in a shared queue, and an idle thread takes its sessions' next task first. The calling side keeps each session's
 * snapshot and journal: a session whose thread was lost, or that left its thread for another one
 * (`maxSessionsPerThread`), is restored on its next call. A task past `hardTimeoutMs` terminates its thread.
 */
import { abortError } from '../async/queue.js'
import type { CellValue, CopyTextOptions, Explanation, ValuesOptions } from '../contract.js'
import type { AbortSignalLike, AsyncStructureResult, AsyncSession, CachedValues, Calc, CalcLimits, CalcOptions, CalcPoolOptions, ChangingMethod, Clock, ComputeResult, ExplainOptions, SessionHistory, SessionOptions, UndoResult } from '../contract.js'
import { CalcError } from '../exec/budget.js'
import { InputError } from '../input-error.js'
import { addedLocales, onLocale } from '../locale/registry.js'
import { onRegexLevel, regexLevel } from '../regex-level.js'
import { splitCall, checkCall, checkEditCall, withCall } from '../async/options.js'
import { checkedMapCell } from '../structure/index.js'
import { subscribers, withSubscribe } from '../session/subscribers.js'
import type { SentGeometry } from '../structure/geometry.js'
import { FILL_OPTIONS, FIND_CALL_OPTIONS, READ_OPTIONS, REPLACE_OPTIONS, strictMethods, TEXT_OPTIONS } from '../strict-args.js'
import type { Spawn, Thread, ThreadLoss } from './platform.js'
import { texts, unpackSlowly, type Texts } from './pack.js'
import { CHANGING, fromData, sentLocale, type Answer, type Entry, type FromThread, type Op, type Sent, type Snapshot, type ToThread } from './protocol.js'

interface Timers {
  setTimeout(f: () => void, ms: number): unknown
  clearTimeout(t: unknown): void
  structuredClone<T>(v: T): T
}
const g = globalThis as unknown as Timers

type Message = ToThread extends infer M ? (M extends { task: number } ? Omit<M, 'task'> : never) : never

interface Task {
  id: number
  msg: Message
  session?: Rec
  /** The pool's own (restore, snapshot, leaving a thread): not counted in `maxQueue`. */
  internal?: boolean
  signal?: AbortSignalLike
  /** Its book is still being written as text: it waits at the head of its queue, and so do those after it. */
  pending?: boolean
  unlisten?: () => void
  resolve(v: unknown): void
  reject(e: unknown): void
}

interface Slot {
  /** None after a loss, until a task needs one: a thread that cannot start fails one task at a time. */
  thread: Thread | null
  busy: Task | null
  timer: unknown
  /** Tasks of its sessions, in order. */
  queue: Task[]
  /** Sessions placed on the thread, and those it holds now. */
  sessions: Set<Rec>
  resident: Set<Rec>
}

interface Rec {
  sid: number
  slot: Slot
  snapshot: Snapshot
  journal: Entry[]
  /** The last changing call answered. */
  seq: number
  clock: Clock
  history: SessionHistory
  revision: number
  /** When it was last called, for leaving a full thread. */
  used: number
  /** Closed, or its restore failed: every call is refused with it. */
  gone: CalcError | null
  snapshotting: boolean
}

/** The argument a task carries to its thread, as an `InputError` names it. */
const ARGUMENT: Record<string, string> = { compute: 'book', open: 'book', setCells: 'edits', replaceCells: 'edits', clearCells: 'ranges', patchCells: 'patches', setCellText: 'text', setCellTexts: 'entries', previewCellTexts: 'entries', pasteText: 'text', previewPasteText: 'text', fillCellText: 'text', previewFillCellText: 'text', setLayout: 'edits', setNames: 'names', setExternalBooks: 'books', setClock: 'clock', editStructure: 'op', setStyles: 'styles', explain: 'opts', getViewport: 'opts', getResult: 'opts', getSnapshot: 'opts' }

/** What a structured clone refused: input outside the contract (a function in `meta`, say), refused as the synchronous engine would refuse a wrong type. */
const unsendable = (what: string, e: unknown): InputError => new InputError(ARGUMENT[what] ?? what, `cannot be sent to a thread: ${e instanceof Error ? e.message : String(e)} (values and meta must survive a structured clone)`)


/** What tests look into: the threads and the sessions. */
export interface PoolInside {
  slots: Slot[]
  recs: Set<Rec>
}

/** Bytes of a thread's heap per value of the default `limits.cells`: twice the worst measured for one array that took the process down (`docs/async.md`). */
const HEAP_PER_CELL = 64
/** Bytes of a thread's heap per value of the default `limits.spill`: over three times the worst measured for one spill that took the process down. */
const HEAP_PER_SPILLED = 1024
/** The part of a thread's heap the range caches of its books may hold together (`docs/cache.md`). */
const CACHE_SHARE = 1 / 4

/** The part of a thread's heap the histories of its sessions (`Session.undo`) hold together by default. */
const UNDO_SHARE = 1 / 8

/**
 * The budget of a session's history in a thread of `maxHeapMb` where the session gives none: its share of an eighth of
 * the heap among the sessions a thread holds (`maxSessionsPerThread`, else the pool's sessions over its threads, else
 * one), in whole megabytes, one at least. Without `maxHeapMb`, the session's own default.
 */
function undoMb(o: CalcPoolOptions): number | undefined {
  if (!o.maxHeapMb) return undefined
  const sessions = o.maxSessionsPerThread ?? (o.maxSessions ? Math.ceil(o.maxSessions / (o.threads ?? 2)) : 1)
  return Math.max(1, Math.floor((o.maxHeapMb * UNDO_SHARE) / sessions))
}

/** The limits of a pool's computations: with `maxHeapMb`, `cells` and `spill` by the heap where the host gives none. */
function poolLimits(o: CalcPoolOptions): CalcLimits | undefined {
  if (!o.maxHeapMb) return o.limits
  const heap = o.maxHeapMb * 2 ** 20
  const cells = o.limits?.cells ?? Math.floor(heap / HEAP_PER_CELL)
  // a host that raises `cells` keeps the spill's bound by the heap
  return { ...o.limits, cells, spill: o.limits?.spill ?? Math.min(cells, Math.floor(heap / HEAP_PER_SPILLED)) }
}

export function createPool(o: CalcPoolOptions, spawn: Spawn, peek?: (inside: PoolInside) => void): Calc {
  const perThread = o.maxSessionsPerThread ?? Infinity
  const snapshotEvery = o.snapshotEvery ?? 64
  const limits = poolLimits(o)
  const historyMb = undoMb(o)
  const jobs: Task[] = []
  const recs = new Set<Rec>()
  let ids = 0
  let clock = 0
  let closed: CalcError | null = null

  const start = (slot: Slot): Thread => {
    const t: Thread = spawn(
      (m) => slot.thread === t && answered(slot, m),
      (l) => slot.thread === t && lose(slot, l),
    )
    // a thread computes in the locales loaded here, those loaded later as they are
    const locales = addedLocales().map(sentLocale)
    t.post({ t: 'init', sliceMs: o.sliceMs ?? 50, ...(limits ? { limits } : {}), ...(o.maxHeapMb ? { cacheBytes: o.maxHeapMb * 2 ** 20 * CACHE_SHARE } : {}), ...(locales.length ? { locales } : {}), ...(regexLevel() ? { regex: regexLevel()!.kind } : {}) })
    return t
  }
  const slots: Slot[] = Array.from({ length: o.threads ?? 2 }, () => {
    const slot = { busy: null, timer: undefined, queue: [], sessions: new Set(), resident: new Set() } as unknown as Slot
    slot.thread = start(slot)
    return slot
  })
  const unlistenLocales = onLocale((l) => {
    for (const s of slots) s.thread?.post({ t: 'locales', locales: [sentLocale(l)] })
  })
  const unlistenRegex = onRegexLevel(() => {
    for (const s of slots) s.thread?.post({ t: 'regex', kind: regexLevel()!.kind })
  })

  peek?.({ slots, recs })
  const waiting = () => jobs.filter((t) => !t.internal).length + slots.reduce((n, s) => n + s.queue.filter((t) => !t.internal).length, 0)

  function finish(task: Task) {
    task.unlisten?.()
    task.unlisten = undefined
  }
  const fail = (task: Task, e: unknown) => {
    finish(task)
    task.reject(task.signal?.aborted && e instanceof CalcError && e.code === 'aborted' ? abortError(task.signal) : e)
  }

  /** A task that will not run: the caller's fails with `why`; the pool's own (a snapshot, leaving a thread) have nothing left to do. */
  const refuse = (task: Task, why: unknown) => (task.internal ? task.resolve(null) : fail(task, why))

  /** A session is gone: its waiting tasks are refused. */
  function bury(r: Rec, why: CalcError) {
    r.gone ??= why
    recs.delete(r)
    r.slot.sessions.delete(r)
    r.slot.resident.delete(r)
    for (const t of r.slot.queue.filter((x) => x.session === r && x.msg.t !== 'drop')) {
      r.slot.queue.splice(r.slot.queue.indexOf(t), 1)
      refuse(t, why)
    }
  }

  /** A task that did not reach its thread (`e`, why): it fails alone, and pump does not build the same restore again. */
  function unsent(slot: Slot, task: Task, e: unknown) {
    const m = task.msg
    const r = task.session
    if (!slot.thread) {
      const why = new CalcError('thread_failed', `the thread could not start: ${e instanceof Error ? e.message : String(e)}`, { cause: e })
      // a restore's thread did not start: the call waiting on it fails, and the session comes back on its next call
      if (r && m.t === 'restore' && slot.queue[0]?.session === r) fail(slot.queue.shift()!, why)
      return refuse(task, why)
    }
    const why = unsendable(m.t === 'call' ? m.op : m.t, e)
    // a restore that cannot be sent closes its session, as a restore that fails does
    if (r && m.t === 'restore') bury(r, new CalcError('restore_failed', `the session could not be restored: ${why.message}`, { cause: why }))
    refuse(task, why)
  }

  function send(slot: Slot, task: Task) {
    // a changing call gets the session's next sequence number when it reaches the thread
    const m = task.msg
    if (m.t === 'call' && task.session && (CHANGING as readonly string[]).includes(m.op)) task.msg = { ...m, seq: task.session.seq + 1 }
    try {
      slot.thread ??= start(slot)
      slot.thread.post({ ...task.msg, task: task.id } as ToThread)
    } catch (e) {
      // nothing reached the thread: the task fails alone, the thread stays free
      return unsent(slot, task, e)
    }
    slot.busy = task
    slot.thread.hold(true)
    if (o.hardTimeoutMs === undefined) return
    // a restore opens the book and replays the journal: each of those was answered within the deadline once
    const ms = o.hardTimeoutMs * (task.msg.t === 'restore' ? 1 + task.msg.journal.length : 1)
    slot.timer = g.setTimeout(() => slot.busy === task && lose(slot, { code: 'hard_timeout', message: `the task ran past its hard deadline (${ms} ms); its thread was terminated` }), ms)
  }

  /** The thread's next task: its sessions' first, then a one-off computation. */
  function pump(slot: Slot) {
    while (!slot.busy && !closed) {
      const head = slot.queue[0] ?? jobs[0]
      if (!head) {
        slot.thread?.hold(false)
        return
      }
      if (head.pending) return
      let task = head === slot.queue[0] ? slot.queue.shift()! : jobs.shift()!
      const r = task.session
      const t = task.msg.t
      if (r && !slot.resident.has(r) && t !== 'open' && t !== 'restore') {
        // leaving or snapshotting a thread it is not on needs nothing; any other task brings it back first
        if (t === 'drop' || t === 'snapshot') {
          if (t === 'snapshot') r.snapshotting = false
          task.resolve(null)
          continue
        }
        slot.queue.unshift(task)
        task = { id: ++ids, msg: { t: 'restore', sid: r.sid, snapshot: r.snapshot, journal: r.journal }, session: r, internal: true, resolve() {}, reject() {} }
      }
      if ((task.msg.t === 'open' || task.msg.t === 'restore') && slot.resident.size >= perThread) {
        const out = [...slot.resident].filter((x) => x !== r).sort((a, b) => a.used - b.used)[0]
        if (out) {
          slot.queue.unshift(task)
          slot.resident.delete(out)
          // it leaves with a snapshot, so that bringing it back replays nothing
          task = { id: ++ids, msg: { t: 'drop', sid: out.sid, keep: out.journal.length > 0 }, session: out, internal: true, resolve() {}, reject() {} }
        }
      }
      send(slot, task)
    }
  }

  function answered(slot: Slot, m: FromThread) {
    const task = slot.busy
    if (!task || task.id !== m.task) return
    g.clearTimeout(slot.timer)
    slot.busy = null
    const r = task.session
    if (m.ok) settled(slot, task, m.value)
    else {
      const e = fromData(m.error)
      if (r && task.msg.t === 'restore') bury(r, new CalcError('restore_failed', `the session could not be restored: ${e.message}`, { cause: e }))
      if (task.msg.t === 'snapshot' && r) r.snapshotting = false
      fail(task, e)
    }
    pump(slot)
  }

  function settled(slot: Slot, task: Task, value: unknown) {
    const r = task.session
    finish(task)
    const msg = task.msg
    if (!r) return task.resolve(value)
    switch (msg.t) {
      case 'open':
      case 'restore':
        slot.resident.add(r)
        if (msg.t === 'open') {
          r.snapshot.clock = r.clock = (value as Answer).clock
          r.history = (value as Answer).history
          r.revision = (value as Answer).revision
        }
        return task.resolve(value)
      case 'snapshot':
      case 'drop': {
        const s = value as { book: Sent; clock: Clock; seq: number; history: Snapshot['history']; revision: number } | null
        if (s) {
          r.snapshot = { book: s.book, opts: r.snapshot.opts, clock: s.clock, seq: s.seq, history: s.history, revision: s.revision }
          r.journal = r.journal.filter((e) => e.seq > s.seq)
        }
        if (msg.t === 'snapshot') r.snapshotting = false
        return task.resolve(null)
      }
      case 'call': {
        const a = value as Answer
        r.clock = a.clock
        r.history = a.history
        r.revision = a.revision
        if (msg.seq !== undefined && msg.seq > r.seq) {
          r.seq = msg.seq
          r.journal.push({ seq: msg.seq, op: msg.op, args: msg.args, clock: a.clock })
          if (r.journal.length >= snapshotEvery && !r.snapshotting) {
            r.snapshotting = true
            // next on the thread, ahead of calls already waiting, so that pipelined calls do not grow the journal past it
            slot.queue.unshift({ id: ++ids, msg: { t: 'snapshot', sid: r.sid }, session: r, internal: true, resolve() {}, reject() {} })
          }
        }
        return task.resolve(a)
      }
      default:
        return task.resolve(value)
    }
  }

  /** The thread is gone (terminated past its deadline, out of memory, failed): its task fails, its sessions come back on their next call, on a new thread. */
  function lose(slot: Slot, l: ThreadLoss | { code: 'hard_timeout'; message: string }) {
    const task = slot.busy
    g.clearTimeout(slot.timer)
    slot.thread?.terminate()
    slot.thread = null
    slot.busy = null
    slot.resident.clear()
    if (task) {
      const e = new CalcError(l.code, l.message)
      const r = task.session
      if (r && task.msg.t === 'restore') bury(r, new CalcError('restore_failed', `the session could not be restored: ${l.message}`, { cause: e }))
      if (task.msg.t === 'snapshot' && r) r.snapshotting = false
      // a session leaving its thread has left it when the thread is gone
      if (task.msg.t === 'drop') task.resolve(null)
      else fail(task, e)
    }
    pump(slot)
  }

  /** A task in the queue of `slot` (a session's) or in the shared one; `book`, its book and stored values while they are written as text. */
  function submit(msg: Message, at?: { slot: Slot; session: Rec }, signal?: AbortSignalLike, internal = false, book?: Promise<Texts>): Promise<unknown> {
    // a task refused here leaves its text being written: a failure of it is not left unhandled
    void book?.catch(() => {})
    if (closed) return Promise.reject(closed)
    if (at?.session.gone) return Promise.reject(at.session.gone)
    if (signal?.aborted) return Promise.reject(abortError(signal))
    // a task an idle thread takes at once does not wait
    const idle = at ? !at.slot.busy && !at.slot.queue.length : !jobs.length && slots.some((s) => !s.busy && !s.queue.length)
    if (!internal && !idle && o.maxQueue !== undefined && waiting() >= o.maxQueue) return Promise.reject(new CalcError('queue_full', `${o.maxQueue} tasks are already waiting for a thread`))
    return new Promise((resolve, reject) => {
      const task: Task = { id: ++ids, msg, session: at?.session, signal, internal, resolve, reject }
      const queue = at ? at.slot.queue : jobs
      const go = () => (at ? pump(at.slot) : slots.forEach(pump))
      // a book written between turns: the task keeps its place in the queue, and goes when the text is ready
      if (book) {
        task.pending = true
        void book.then(
          (sent) => {
            task.pending = false
            Object.assign(task.msg, sent)
            go()
          },
          (e) => {
            // the text could not be written (a meta that throws when read): the call fails as it does for a small book, and the queue moves on
            const i = queue.indexOf(task)
            if (i >= 0) {
              queue.splice(i, 1)
              fail(task, e)
            }
            go()
          },
        )
      }
      if (signal) {
        const onAbort = () => {
          const i = queue.indexOf(task)
          if (i >= 0) {
            queue.splice(i, 1)
            fail(task, abortError(signal))
          } else {
            const slot = slots.find((s) => s.busy === task)
            slot?.thread?.post({ t: 'cancel', task: task.id })
          }
        }
        signal.addEventListener('abort', onAbort, { once: true })
        task.unlisten = () => signal.removeEventListener('abort', onAbort)
      }
      queue.push(task)
      go()
    })
  }

  function session(r: Rec): AsyncSession {
    const subs = subscribers(r.revision)
    /** A changing call's answer, told to the session's subscribers before the call's promise settles. */
    const told = <T>(method: ChangingMethod, p: Promise<T>) => p.then((v) => subs.after(method, v))
    /** A call to the session's thread; a changing call's arguments, its edit options among them, go into its journal as a copy. */
    const send = (op: Op, args: unknown[], c: CalcOptions): Promise<Answer> => {
      r.used = ++clock
      const changing = (CHANGING as readonly string[]).includes(op)
      let kept = args
      try {
        if (changing) kept = g.structuredClone(args)
      } catch (e) {
        throw unsendable(op, e)
      }
      const msg: Message = { t: 'call', sid: r.sid, op, args: kept, timeoutMs: c.timeoutMs ?? o.timeoutMs }
      return submit(msg, { slot: r.slot, session: r }, c.signal) as Promise<Answer>
    }
    const call = (op: Op, args: unknown[], co?: unknown): Promise<Answer> => {
      try {
        return send(op, args, checkCall(co))
      } catch (e) {
        return Promise.reject(e)
      }
    }
    /** A call that takes edit options (`EDITS`): they go to the thread after its arguments. */
    const edit = (op: Op, args: unknown[], co?: unknown, extra?: readonly string[]): Promise<Answer> => {
      try {
        const [c, eo] = checkEditCall(co, extra)
        return send(op, eo ? [...args, eo] : args, c)
      } catch (e) {
        return Promise.reject(e)
      }
    }
    const value = <T>(p: Promise<Answer>) => p.then((a) => a.value as T)
    /** A structure operation's answer, done, undone or redone: its mapCell made here from the geometry the thread sent. */
    const mapped = <T>(p: Promise<Answer>) =>
      value<(T & { geometry?: SentGeometry }) | null>(p).then((v): T | null => {
        if (!v?.geometry) return v
        const { geometry: g, ...rest } = v
        return { ...rest, mapCell: checkedMapCell(g) } as T
      })
    const step = (op: 'undo' | 'redo', co: unknown) => mapped<UndoResult>(edit(op, [], co))
    return withSubscribe(strictMethods<AsyncSession>({
      get clock() {
        return { ...r.clock }
      },
      get history() {
        return { ...r.history }
      },
      get revision() {
        return r.revision
      },
      applyChanges: (ops, co) => told('applyChanges', value(edit('applyChanges', [ops], co, ['clock']))),
      previewChanges: (ops, co) => value(edit('previewChanges', [ops], co, ['clock', 'maxChanges'])),
      readCells: (refs, co) => value(edit('readCells', [refs], co, READ_OPTIONS)),
      findCells: (refs, query, co) => value(edit('findCells', [refs, query], co, FIND_CALL_OPTIONS)),
      replaceInCells: (refs, query, replacement, co) => told('replaceInCells', value(edit('replaceInCells', [refs, query, replacement], co, REPLACE_OPTIONS))),
      getDependents: (refs, co) => value(edit('getDependents', [refs], co, ['maxCells', 'cursor', 'count', 'maxReferences'])),
      summarizeRanges: (refs, co) => value(edit('summarizeRanges', [refs], co, ['rows'])),
      setCells: (edits, co) => told('setCells', value(edit('setCells', [edits], co))),
      replaceCells: (edits, co) => told('replaceCells', value(edit('replaceCells', [edits], co))),
      clearCells: (ranges, what, co) => told('clearCells', value(edit('clearCells', [ranges, what], co))),
      patchCells: (patches, co) => told('patchCells', value(edit('patchCells', [patches], co))),
      setCellText: (sheet, cell, text, co) => told('setCellText', value(edit('setCellText', [sheet, cell, text], co, TEXT_OPTIONS))),
      previewCellText: (sheet, cell, text, co) => value(edit('previewCellText', [sheet, cell, text], co, TEXT_OPTIONS)),
      setCellTexts: (entries, co) => told('setCellTexts', value(edit('setCellTexts', [entries], co, TEXT_OPTIONS))),
      previewCellTexts: (entries, co) => value(edit('previewCellTexts', [entries], co, TEXT_OPTIONS)),
      pasteText: (sheet, cell, text, co) => told('pasteText', value(edit('pasteText', [sheet, cell, text], co))),
      previewPasteText: (sheet, cell, text, co) => value(edit('previewPasteText', [sheet, cell, text], co)),
      fillCellText: (sheet, range, text, co) => told('fillCellText', value(edit('fillCellText', [sheet, range, text], co, FILL_OPTIONS))),
      previewFillCellText: (sheet, range, text, co) => value(edit('previewFillCellText', [sheet, range, text], co, FILL_OPTIONS)),
      setNames: (names, co) => told('setNames', value(edit('setNames', [names], co))),
      setExternalBooks: (books, co) => told('setExternalBooks', value(edit('setExternalBooks', [books], co))),
      setClock: (next, co) => told('setClock', value(edit('setClock', [next], co))),
      editStructure: (op, co) => told('editStructure', mapped<AsyncStructureResult>(edit('editStructure', [op], co)) as Promise<AsyncStructureResult>),
      setStyles: (styles, co) => told('setStyles', value(edit('setStyles', [styles], co))),
      setLayout: (edits, co) => told('setLayout', value(edit('setLayout', [edits], co))),
      getBook: () => value<Sent>(call('getBook', [])).then(unpackSlowly),
      getLayout: () => value(call('getLayout', [])),
      getUsedRange: (sheet) => value(call('getUsedRange', [sheet])),
      getAutoFilterRange: (sheet) => value(call('getAutoFilterRange', [sheet])),
      getEnd: (sheet, cell, direction) => value(call('getEnd', [sheet, cell, direction])),
      getAutoSum: (sheet, cell) => value(call('getAutoSum', [sheet, cell])),
      getNamesFromSelection: (sheet, range, labels) => value(call('getNamesFromSelection', [sheet, range, labels])),
      getSpecialCells: (sheet, range, kind) => value(call('getSpecialCells', [sheet, range, kind])),
      explain: (sheet, cell, eo) => withCall<ExplainOptions, Explanation>(eo, (rest, co) => value(call('explain', [sheet, cell, rest], co))),
      getCells: (sheet, range, co) => value(call('getCells', [sheet, range], co)),
      copyText: (sheet, range, o) => withCall<CopyTextOptions, string>(o, (rest, co) => value(call('copyText', [sheet, range, rest], co))),
      getValues: (sheet, range, o) => withCall<ValuesOptions, CellValue[][]>(o, (rest, co) => value(call('getValues', [sheet, range, rest], co))),
      getCell: (sheet, cell) => value(call('getCell', [sheet, cell])),
      getViewport: (sheet, range, vo) => value(call('getViewport', [sheet, range, vo])),
      getCellIssue: (sheet, cell) => value(call('getCellIssue', [sheet, cell])),
      getValidation: (sheet, cell) => value(call('getValidation', [sheet, cell])),
      validateEntry: (sheet, cell, input, co) => value(call('validateEntry', [sheet, cell, input], co)),
      getResult: (co) => value(call('getResult', [], co)),
      // the book comes as text, read a chunk at a time as `getBook()` reads it
      getSnapshot: (co) => value<{ book: Sent; result: ComputeResult }>(call('getSnapshot', [], co)).then(async (v) => ({ book: await unpackSlowly(v.book), result: v.result })),
      getHash: (co) => value(call('getHash', [], co)),
      undo: (co) => told('undo', step('undo', co)),
      redo: (co) => told('redo', step('redo', co)),
      subscribe: (listener) => subs.subscribe(listener),
      async close() {
        subs.clear()
        if (r.gone) return
        const done = submit({ t: 'drop', sid: r.sid }, { slot: r.slot, session: r }, undefined, true)
        r.gone = new CalcError('closed', 'the session is closed')
        recs.delete(r)
        r.slot.sessions.delete(r)
        await done
        r.slot.resident.delete(r)
      },
    }, true), subs)
  }

  return strictMethods<Calc>({
    async compute(book, copts) {
      const { signal, timeoutMs, rest } = splitCall<{ cached?: CachedValues }>(copts)
      const { cached, ...opts } = rest ?? {}
      const [sent, later] = texts(book, cached)
      return (await submit({ t: 'compute', ...sent, opts, timeoutMs: timeoutMs ?? o.timeoutMs }, undefined, signal, false, later)) as never
    },
    async createSession(book, sopts) {
      if (closed) throw closed
      if (o.maxSessions !== undefined && recs.size >= o.maxSessions) throw new CalcError('too_many_sessions', `the pool already holds ${o.maxSessions} sessions`)
      const { signal, timeoutMs, rest } = splitCall<SessionOptions>(sopts)
      const { clock: given, cached, ...own } = { ...rest } as SessionOptions & { clock?: Clock }
      // the session's own copy of its options, as its snapshot keeps them: what the caller changes in them later (its
      // externals' values, say) does not reach a restore
      let opts: Omit<SessionOptions, 'cached'>
      try {
        opts = g.structuredClone(own)
      } catch (e) {
        throw unsendable('opts', e)
      }
      // a session without a budget of its own for its history takes its share of the thread's heap; wrong input stays for the thread to refuse
      const u = opts.undo as unknown
      if (historyMb !== undefined && (u === undefined || (typeof u === 'object' && u !== null && (u as { maxMb?: unknown }).maxMb === undefined))) opts.undo = { ...(u as object), maxMb: historyMb }
      const slot = [...slots].sort((a, b) => a.sessions.size - b.sessions.size || a.queue.length - b.queue.length)[0]
      // the book and the stored values are taken now and written as text, between turns of the event loop when they are large; the text is the first snapshot
      const [sent, written] = texts(book, cached)
      const later = written?.then((x) => {
        Object.assign(r.snapshot, x)
        return x
      })
      const r: Rec = { sid: ++ids, slot, snapshot: { ...sent, opts, clock: given!, seq: 0 }, journal: [], seq: 0, clock: given!, history: { undo: 0, redo: 0 }, revision: 0, used: ++clock, gone: null, snapshotting: false }
      recs.add(r)
      slot.sessions.add(r)
      await submit({ t: 'open', sid: r.sid, book: r.snapshot.book, ...(sent.cached ? { cached: sent.cached } : {}), opts: { ...opts, ...(given ? { clock: given } : {}) }, timeoutMs: timeoutMs ?? o.timeoutMs }, { slot, session: r }, signal, false, later).catch((e) => {
        bury(r, new CalcError('closed', 'the session did not open'))
        throw e
      })
      return session(r)
    },
    async close() {
      if (closed) return
      closed = new CalcError('closed', 'the pool is closed')
      unlistenLocales()
      unlistenRegex()
      for (const t of jobs.splice(0)) fail(t, closed)
      for (const s of slots) {
        // a session closing is closed with the pool; the pool's own tasks end
        for (const t of s.queue.splice(0)) (t.msg.t === 'drop' ? t.resolve(null) : refuse(t, closed))
        g.clearTimeout(s.timer)
        if (s.busy) (s.busy.msg.t === 'drop' ? s.busy.resolve(null) : refuse(s.busy, closed))
        s.busy = null
        s.thread?.terminate()
      }
      for (const r of recs) r.gone = closed
      recs.clear()
    },
  })
}
