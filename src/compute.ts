/** A whole book computed, in steps (`exec/budget.ts`): `compute` runs them at once, the asynchronous calls in portions. */
import type { BookInput, ComputeOptions, ComputeResult } from './contract.js'
import { buildState } from './engine.js'
import type { Steps } from './exec/budget.js'
import { currentClock, nowMs } from './env.js'
import { checkBook } from './input.js'
import { buildResult, clockSerial } from './result.js'

/** The computation takes the book's cells as they are: nothing writes into them (a held value copies its sheet, `held.ts`). */
const owned = () => true

/**
 * `o` checked already (`checkComputeOptions`). `hosted`: the host runs between the steps (`createInlineCalc`) and may
 * change the book meanwhile, so the computation copies its cells first.
 */
export function* computeSteps(book: BookInput, o: ComputeOptions, hosted = false): Steps<ComputeResult> {
  const t0 = nowMs()
  book = checkBook(book)
  const clock = o.clock ?? currentClock()
  const state = yield* buildState(book, { now: clockSerial(clock), seed: clock.seed, externals: o.externals, execution: o.execution, cached: o.cached, own: hosted ? undefined : owned })
  return yield* buildResult(state, clock, t0, o.display !== false)
}
