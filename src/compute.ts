/** A whole book computed, in steps (`exec/budget.ts`): `compute` runs them at once, the asynchronous calls in portions. */
import type { BookInput, ComputeOptions, ComputeResult } from './contract.js'
import { buildState } from './engine.js'
import type { Steps } from './exec/budget.js'
import { currentClock, nowMs } from './env.js'
import { checkBook } from './input.js'
import { buildResult, clockSerial } from './result.js'

/** `o` checked already (`checkComputeOptions`). */
export function* computeSteps(book: BookInput, o: ComputeOptions): Steps<ComputeResult> {
  const t0 = nowMs()
  checkBook(book)
  const clock = o.clock ?? currentClock()
  const state = yield* buildState(book, { now: clockSerial(clock), seed: clock.seed, externals: o.externals, execution: o.execution })
  return yield* buildResult(state, clock, t0, o.display !== false)
}
