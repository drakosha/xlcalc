/**
 * Runs steps in portions (`docs/async.md`): each portion with the call's budget active, then a yield to the event
 * loop. An abort or a deadline seen while yielding is thrown into the steps where they stopped, so a session's call
 * puts back what it changed.
 */
import { Budget, CalcError, portion, type Steps } from '../exec/budget.js'
import type { AbortSignalLike, CalcLimits } from '../contract.js'
import { nowMs } from '../env.js'

interface Timers {
  setImmediate?: (f: () => void) => unknown
  setTimeout?: (f: () => void, ms: number) => unknown
  clearTimeout?: (t: unknown) => void
  MessageChannel?: new () => { port1: { onmessage: (() => void) | null }; port2: { postMessage(m: unknown): void } }
}

const g = globalThis as unknown as Timers

// one channel for the browser's yields: a MessageChannel message is a task that no timer clamping delays
let channel: { post(): void; waiting: (() => void)[] } | undefined
function viaChannel(): Promise<void> {
  if (!channel) {
    const ch = new g.MessageChannel!()
    const waiting: (() => void)[] = []
    ch.port1.onmessage = () => waiting.shift()?.()
    channel = { post: () => ch.port2.postMessage(null), waiting }
  }
  const c = channel
  return new Promise((resolve) => {
    c.waiting.push(resolve)
    c.post()
  })
}

/**
 * Gives way to the event loop: `setImmediate` in Node, a message in a browser, a timer elsewhere. Not
 * `scheduler.yield`: its continuation goes before the page's own tasks, timers and workers' messages starve.
 */
export function pause(): Promise<void> {
  if (g.setImmediate) return new Promise((resolve) => g.setImmediate!(resolve))
  if (g.MessageChannel) return viaChannel()
  if (g.setTimeout) return new Promise((resolve) => g.setTimeout!(resolve, 0))
  return Promise.resolve()
}

/** A call's budget: its limits, its signal, and its slices of `sliceMs`; the deadline is set when it starts computing. */
export const callBudget = (limits: CalcLimits | undefined, signal: AbortSignalLike | undefined, sliceMs: number): Budget => new Budget(limits, Infinity, signal, sliceMs)

/** `p`, or nothing past `ms` or once `signal` aborts: the cells still waiting for answers from outside the workbook stay so. */
function within(p: Promise<void>, ms: number, signal: AbortSignalLike | undefined): Promise<void> {
  if (ms <= 0 || signal?.aborted) return Promise.resolve()
  return new Promise((resolve) => {
    const done = () => {
      if (timer !== undefined) g.clearTimeout?.(timer)
      signal?.removeEventListener('abort', done)
      resolve()
    }
    const timer = Number.isFinite(ms) && g.setTimeout ? g.setTimeout(done, ms) : undefined
    signal?.addEventListener('abort', done, { once: true })
    p.then(done, done)
  })
}

/** The steps to their end, in portions; `timeoutMs` counts from now. */
export async function drive<T>(steps: Steps<T>, b: Budget, timeoutMs?: number): Promise<T> {
  if (timeoutMs !== undefined) b.deadline = nowMs() + timeoutMs
  b.driven = true
  b.waitUntil = b.deadline
  const early = b.stopped()
  if (early) {
    steps.return(undefined as T)
    throw early
  }
  let r = portion(steps, b)
  while (!r.done) {
    const w = b.waiting
    if (w) {
      // answers are waited for until the call's deadline as it started, which then moves on by the wait: it bounds the computing alone
      b.waiting = undefined
      const t0 = nowMs()
      await within(w, b.waitUntil - t0, b.signal)
      b.deadline += nowMs() - t0
    } else await pause()
    const stop: CalcError | null = b.stopped()
    r = stop ? portion(steps, b, 'throw', stop) : portion(steps, b)
  }
  return r.value
}
