/**
 * A workbook formula engine (the package root). The contract (types only) is also available as the
 * `./contract` subpath; the implementation compiles formulas to JS and computes over the reference graph.
 * The `./unstable/*` subpaths (internals, the synthetic book) are outside semver.
 */
import type { BookInput, ComputeOptions, ComputeResult, FormulaEngine, NameCheck, ValidateContext, ValidateResult } from './contract.js'
import { openSession } from './session.js'
import { createState, debugSource } from './engine.js'
import { currentClock, nowMs } from './env.js'
import { checkBook, checkComputeOptions, checkString, checkValidateContext } from './input.js'
import { buildResult, clockSerial } from './result.js'
import { validate, validateName as checkName } from './syntax.js'

export * from './contract.js'
export { FUNCTION_NAMES } from './functions.js'
export { CONFIG_VERSION } from './config.js'
export { openSession } from './session.js'
export { InputError } from './input.js'
export { enUS, ruRU, toInvariant, toLocal, validateLocalFormula, type Locale, type Translation } from './locale/index.js'

/**
 * Computes the whole book. Without `opts.clock` the clock is the current moment and a random seed; the
 * result records it in `clock`, and passing that clock again reproduces the result. Input of the wrong
 * shape throws `InputError`; invalid formulas do not throw, they give error values with reasons in `issues`.
 */
export function compute(book: BookInput, opts?: ComputeOptions): ComputeResult {
  const t0 = nowMs()
  const o = checkComputeOptions(opts)
  checkBook(book)
  const clock = o.clock ?? currentClock()
  const state = createState(book, { now: clockSerial(clock), seed: clock.seed, externals: o.externals })
  return buildResult(state, clock, t0)
}

export function validateFormula(text: string, ctx?: ValidateContext): ValidateResult {
  const r = validate(checkString(text, 'text'), checkValidateContext(ctx))
  if (!r.ok) return r
  const { ast: _ast, ...rest } = r
  return rest
}

export function validateName(name: string): NameCheck {
  return checkName(checkString(name, 'name'))
}

/** Generated code of the book in evaluation order, for reading while debugging; not part of the contract. */
export function debugGeneratedCode(book: BookInput, opts?: ComputeOptions): string {
  const o = checkComputeOptions(opts)
  checkBook(book)
  const clock = o.clock ?? currentClock()
  return debugSource(createState(book, { now: clockSerial(clock), seed: clock.seed, externals: o.externals }))
}

/** The contract implementation as one object. */
export const engine: FormulaEngine = { compute, validateFormula, validateName, openSession }
