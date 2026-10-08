/**
 * The pool of threads (`docs/async.md`). A session is placed on a thread and stays there; one-off computations wait
 * in a shared queue, and an idle thread takes its sessions' next task first. The calling side keeps each session's
 * snapshot and journal: a session whose thread was lost, or that left its thread for another one
 * (`maxSessionsPerThread`), is restored on its next call. A task past `hardTimeoutMs` terminates its thread.
 */
import { abortError } from '../async/queue.js'
import type { AbortSignalLike, AsyncRestructureResult, AsyncSession, BookInput, Calc, CalcPoolOptions, Clock, SessionOptions, StructureOp } from '../contract.js'
import { CalcError } from '../exec/budget.js'
import { splitCall, checkCall } from '../async/options.js'
import { applyOp, checkedMapCell } from '../structure/index.js'
import type { Spawn, Thread, ThreadLoss } from './platform.js'
import { CHANGING, fromData, type Answer, type Entry, type FromThread, type Op, type Snapshot, type ToThread } from './protocol.js'

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
  /** When it was last called, for leaving a full thread. */
  used: number
  /** Closed, or its restore failed: every call is refused with it. */
  gone: CalcError | null
  snapshotting: boolean
}

/** A book's sheets and cells copied, as a session copies them, so the caller's later changes do not reach the snapshot. */
const copyBook = (b: BookInput): BookInput => ({ ...b, sheets: b.sheets.map((s) => ({ ...s, cells: { ...s.cells } })) })

/** What tests look into: the threads and the sessions. */
export interface PoolInside {
  slots: Slot[]
  recs: Set<Rec>
}

export function createPool(o: CalcPoolOptions, spawn: Spawn, peek?: (inside: PoolInside) => void): Calc {
  const perThread = o.maxSessionsPerThread ?? Infinity
  const snapshotEvery = o.snapshotEvery ?? 64
  const limits = o.maxHeapMb && o.limits?.cells === undefined ? { ...o.limits, cells: Math.floor((o.maxHeapMb * 2 ** 20) / 32) } : o.limits
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
    t.post({ t: 'init', sliceMs: o.sliceMs ?? 50, ...(limits ? { limits } : {}) })
    return t
  }
  const slots: Slot[] = Array.from({ length: o.threads ?? 2 }, () => {
    const slot = { busy: null, timer: undefined, queue: [], sessions: new Set(), resident: new Set() } as unknown as Slot
    slot.thread = start(slot)
    return slot
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

  /** A session is gone: its waiting tasks are refused. */
  function bury(r: Rec, why: CalcError) {
    r.gone ??= why
    recs.delete(r)
    r.slot.sessions.delete(r)
    r.slot.resident.delete(r)
    for (const t of r.slot.queue.filter((x) => x.session === r && x.msg.t !== 'drop')) {
      r.slot.queue.splice(r.slot.queue.indexOf(t), 1)
      if (!t.internal) fail(t, why)
    }
  }

  function send(slot: Slot, task: Task) {
    // a changing call gets the session's next sequence number when it reaches the thread
    const m = task.msg
    if (m.t === 'call' && task.session && (CHANGING as readonly string[]).includes(m.op)) task.msg = { ...m, seq: task.session.seq + 1 }
    slot.busy = task
    slot.thread ??= start(slot)
    slot.thread.hold(true)
    slot.thread.post({ ...task.msg, task: task.id } as ToThread)
    if (o.hardTimeoutMs !== undefined) slot.timer = g.setTimeout(() => slot.busy === task && lose(slot, { code: 'hard_timeout', message: `the task ran past hardTimeoutMs (${o.hardTimeoutMs} ms); its thread was terminated` }), o.hardTimeoutMs)
  }

  /** The thread's next task: its sessions' first, then a one-off computation. */
  function pump(slot: Slot) {
    while (!slot.busy && !closed) {
      let task = slot.queue.shift() ?? jobs.shift()
      if (!task) {
        slot.thread?.hold(false)
        return
      }
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
          task = { id: ++ids, msg: { t: 'drop', sid: out.sid }, session: out, internal: true, resolve() {}, reject() {} }
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
        if (msg.t === 'open') r.snapshot.clock = r.clock = (value as Answer).clock
        return task.resolve(value)
      case 'snapshot': {
        const s = value as { book: BookInput; clock: Clock; seq: number }
        r.snapshot = { book: s.book, opts: r.snapshot.opts, clock: s.clock, seq: s.seq }
        r.journal = r.journal.filter((e) => e.seq > s.seq)
        r.snapshotting = false
        return task.resolve(null)
      }
      case 'call': {
        const a = value as Answer
        r.clock = a.clock
        if (msg.seq !== undefined && msg.seq > r.seq) {
          r.seq = msg.seq
          r.journal.push({ seq: msg.seq, op: msg.op, args: msg.args, clock: a.clock })
          if (r.journal.length >= snapshotEvery && !r.snapshotting) {
            r.snapshotting = true
            slot.queue.push({ id: ++ids, msg: { t: 'snapshot', sid: r.sid }, session: r, internal: true, resolve() {}, reject() {} })
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
      fail(task, e)
    }
    pump(slot)
  }

  /** A task in the queue of `slot` (a session's) or in the shared one. */
  function submit(msg: Message, at?: { slot: Slot; session: Rec }, signal?: AbortSignalLike, internal = false): Promise<unknown> {
    if (closed) return Promise.reject(closed)
    if (at?.session.gone) return Promise.reject(at.session.gone)
    if (signal?.aborted) return Promise.reject(abortError(signal))
    if (!internal && o.maxQueue !== undefined && waiting() >= o.maxQueue) return Promise.reject(new CalcError('queue_full', `${o.maxQueue} tasks already wait for a thread`))
    return new Promise((resolve, reject) => {
      const task: Task = { id: ++ids, msg, session: at?.session, signal, internal, resolve, reject }
      const queue = at ? at.slot.queue : jobs
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
      if (at) pump(at.slot)
      else for (const s of slots) pump(s)
    })
  }

  function session(r: Rec): AsyncSession {
    const call = (op: Op, args: unknown[], co?: unknown): Promise<Answer> => {
      try {
        const c = checkCall(co)
        r.used = ++clock
        // the journal keeps its own copy of what a changing call was given
        const changing = (CHANGING as readonly string[]).includes(op)
        const msg: Message = { t: 'call', sid: r.sid, op, args: changing ? g.structuredClone(args) : args, timeoutMs: c.timeoutMs ?? o.timeoutMs }
        return submit(msg, { slot: r.slot, session: r }, c.signal) as Promise<Answer>
      } catch (e) {
        return Promise.reject(e)
      }
    }
    const value = <T>(p: Promise<Answer>) => p.then((a) => a.value as T)
    return {
      get clock() {
        return { ...r.clock }
      },
      setCells: (edits, co) => value(call('setCells', [edits], co)),
      setNames: (names, co) => value(call('setNames', [names], co)),
      setClock: (next, co) => value(call('setClock', [next], co)),
      async restructure(op, co): Promise<AsyncRestructureResult> {
        const a = await call('restructure', [op], co)
        const v = a.value as AsyncRestructureResult
        if (!v.ok) return v
        const outline = a.outline!
        const m = applyOp(outline, op as StructureOp, () => [])
        if (!m.ok) throw new Error(`the operation applies to the session's book but not to its outline: ${m.message}`)
        return { ok: true, changes: v.changes, mapCell: checkedMapCell(outline, m.mapCell) }
      },
      setStyles: (styles) => value(call('setStyles', [styles])),
      setLayout: (edits, co) => value(call('setLayout', [edits], co)),
      book: () => value(call('book', [])),
      explain: (sheet, cell, eo) => value(call('explain', [sheet, cell, eo])),
      result: (co) => value(call('result', [], co)),
      async close() {
        if (r.gone) return
        const done = submit({ t: 'drop', sid: r.sid }, { slot: r.slot, session: r }, undefined, true)
        r.gone = new CalcError('closed', 'the session is closed')
        recs.delete(r)
        r.slot.sessions.delete(r)
        await done
        r.slot.resident.delete(r)
      },
    }
  }

  return {
    async compute(book, copts) {
      const { signal, timeoutMs, rest } = splitCall<object>(copts)
      return (await submit({ t: 'compute', book, opts: rest ?? {}, timeoutMs: timeoutMs ?? o.timeoutMs }, undefined, signal)) as never
    },
    async openSession(book, sopts) {
      if (closed) throw closed
      if (o.maxSessions !== undefined && recs.size >= o.maxSessions) throw new CalcError('too_many_sessions', `the pool holds ${o.maxSessions} sessions already`)
      const { signal, timeoutMs, rest } = splitCall<SessionOptions>(sopts)
      const opts = { ...rest } as SessionOptions & { clock?: Clock }
      const slot = [...slots].sort((a, b) => a.sessions.size - b.sessions.size || a.queue.length - b.queue.length)[0]
      const given = opts.clock
      delete opts.clock
      const r: Rec = { sid: ++ids, slot, snapshot: { book: copyBook(book), opts, clock: given!, seq: 0 }, journal: [], seq: 0, clock: given!, used: ++clock, gone: null, snapshotting: false }
      recs.add(r)
      slot.sessions.add(r)
      await submit({ t: 'open', sid: r.sid, book, opts: { ...opts, ...(given ? { clock: given } : {}) }, timeoutMs: timeoutMs ?? o.timeoutMs }, { slot, session: r }, signal).catch((e) => {
        bury(r, new CalcError('closed', 'the session did not open'))
        throw e
      })
      return session(r)
    },
    async close() {
      if (closed) return
      closed = new CalcError('closed', 'the pool is closed')
      for (const t of jobs.splice(0)) fail(t, closed)
      for (const s of slots) {
        for (const t of s.queue.splice(0)) if (!t.internal) fail(t, closed)
        g.clearTimeout(s.timer)
        if (s.busy && !s.busy.internal) fail(s.busy, closed)
        s.busy = null
        s.thread?.terminate()
      }
      for (const r of recs) r.gone = closed
      recs.clear()
    },
  }
}
