/**
 * The asynchronous session in the same thread (`docs/async.md`): the steps of `session.ts` run in portions, its calls
 * one after another. A thread of the pool holds its sessions the same way.
 */
import type { AsyncRestructureResult, AsyncSession, AsyncSessionOptions, CalcLimits, CalcOptions, Clock, SessionOptions } from '../contract.js'
import { CalcError, type Budget, type Steps } from '../exec/budget.js'
import { openSteps, type SessionHooks } from '../session.js'
import { callBudget, drive } from './drive.js'
import { checkCall, splitCall } from './options.js'
import { Serial } from './queue.js'

/** What a calculator gives its sessions. */
export interface InlineDefaults {
  sliceMs: number
  timeoutMs?: number
  limits?: CalcLimits
  /** The budgets of the calls running now, which closing the calculator stops. */
  running?: Set<Budget>
}

export interface InlineSession extends AsyncSession {
  /** Stops the session at once: its running call and the waiting ones reject with `why`. */
  kill(why: CalcError): void
}

/** The calls of `steps` with a budget each, from the calculator's defaults and the call's options. */
export function budgeted(d: InlineDefaults, limits: CalcLimits | undefined) {
  return async <T>(steps: Steps<T>, o: CalcOptions): Promise<T> => {
    const b = callBudget(limits, o.signal, d.sliceMs)
    d.running?.add(b)
    try {
      return await drive(steps, b, o.timeoutMs ?? d.timeoutMs)
    } finally {
      d.running?.delete(b)
    }
  }
}

export async function openInlineSession(book: Parameters<typeof openSteps>[0], opts: AsyncSessionOptions | undefined, d: InlineDefaults, hooks?: SessionHooks): Promise<InlineSession> {
  const { signal, timeoutMs, rest } = splitCall<SessionOptions>(opts)
  const o = { ...rest, limits: rest?.limits ?? d.limits } as SessionOptions
  const go = budgeted(d, o.limits)
  const core = await go(openSteps(book, o, hooks), { signal, timeoutMs })
  const queue = new Serial()
  let clock: Clock = core.clock
  // close() called: new calls are refused; kill(): the waiting ones too, and the running one stops
  let ended: CalcError | null = null
  let killed: CalcError | null = null
  let running: Budget | null = null
  const turn = <T>(task: () => Promise<T>, signal?: CalcOptions['signal']): Promise<T> => {
    if (ended) return Promise.reject(ended)
    return queue.run(async () => {
      if (killed) throw killed
      return task()
    }, signal)
  }
  const call = <T>(steps: () => Steps<T>, co?: CalcOptions): Promise<T> => {
    let c: CalcOptions
    try {
      c = checkCall(co)
    } catch (e) {
      return Promise.reject(e)
    }
    return turn(async () => {
      const b = callBudget(o.limits, c.signal, d.sliceMs)
      running = b
      d.running?.add(b)
      try {
        const r = await drive(steps(), b, c.timeoutMs ?? d.timeoutMs)
        clock = core.clock
        return r
      } finally {
        running = null
        d.running?.delete(b)
      }
    }, c.signal)
  }
  const plain = <T>(f: () => T): Promise<T> => turn(async () => f())
  return {
    get clock() {
      return { ...clock }
    },
    setCells: (edits, co) => call(() => core.setCells(edits), co),
    setNames: (names, co) => call(() => core.setNames(names), co),
    setClock: (next, co) => call(() => core.setClock(next), co),
    restructure: (op, co) =>
      call(function* (): Steps<AsyncRestructureResult> {
        const r = yield* core.restructure(op)
        return r.ok ? { ok: true, mapCell: r.mapCell, changes: r.changes } : r
      }, co),
    setStyles: (styles) => plain(() => core.setStyles(styles)),
    setLayout: (edits, co) => call(() => core.setLayout(edits), co),
    book: () => plain(() => core.book()),
    explain: (sheet, cell, eo) => call(() => core.explain(sheet, cell, eo)),
    cells: (sheet, range) => call(() => core.cells(sheet, range)),
    result: (co) => call(() => core.result(), co),
    close() {
      if (ended) return Promise.resolve()
      // closing waits for the calls made before it, and closes the session whatever stops them
      const done = queue.run(async () => core.close())
      ended = new CalcError('closed', 'the session is closed')
      return done
    },
    kill(why: CalcError) {
      ended ??= why
      killed = why
      if (running) running.halted = why
      core.close()
    },
  }
}
