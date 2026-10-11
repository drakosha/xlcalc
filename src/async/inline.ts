/**
 * `createInlineCalc`: computation in the same thread, in portions that give way to the event loop (`docs/async.md`).
 * Each call has its own budget: the soft deadline, the abort signal, the limits.
 */
import type { Calc, ComputeOptions, InlineCalcOptions } from '../contract.js'
import { computeSteps } from '../compute.js'
import { CalcError, type Budget } from '../exec/budget.js'
import { checkComputeOptions } from '../input.js'
import { callLimits, checkInlineOptions, splitCall } from './options.js'
import { budgeted, openInlineSession, type InlineDefaults, type InlineSession } from './session.js'
import { strictMethods } from '../strict-args.js'

/** Milliseconds of computing between two yields by default: a frame stays smooth. */
export const SLICE_MS = 10

export function createInlineCalc(opts?: InlineCalcOptions): Calc {
  const o = checkInlineOptions(opts)
  const running = new Set<Budget>()
  const d: InlineDefaults = { sliceMs: o.sliceMs ?? SLICE_MS, timeoutMs: o.timeoutMs, limits: o.limits, running }
  const sessions = new Set<InlineSession>()
  let closed: CalcError | null = null
  return strictMethods<Calc>({
    async compute(book, copts) {
      if (closed) throw closed
      const { signal, timeoutMs, rest } = splitCall<ComputeOptions>(copts)
      const c = checkComputeOptions(rest)
      return budgeted(d, callLimits(d.limits, c.limits))(computeSteps(book, c), { signal, timeoutMs })
    },
    async createSession(book, sopts) {
      if (closed) throw closed
      const s = await openInlineSession(book, sopts, d)
      if (closed) {
        s.kill(closed)
        throw closed
      }
      sessions.add(s)
      return s
    },
    async close() {
      closed ??= new CalcError('closed', 'the calculator is closed')
      for (const b of running) b.halted = closed
      for (const s of sessions) s.kill(closed)
      sessions.clear()
    },
  })
}
