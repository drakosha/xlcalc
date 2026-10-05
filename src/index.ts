/**
 * `@mike/formula`: a workbook formula engine. The contract is in `./contract`; the implementation compiles
 * formulas to JS and computes over the reference graph. The synthetic book generator is in `./synth`,
 * internals for benchmark harnesses in `./internal`.
 */
import type { BookInput, ComputeOptions, ComputeResult, FormulaEngine, NameCheck, ValidateContext, ValidateResult } from './contract.js'
import { openSession } from './session.js'
import { createState, debugSource } from './engine.js'
import { nowMs } from './env.js'
import { buildResult, clockSerial } from './result.js'
import { validate, validateName as checkName } from './syntax.js'

export * from './contract.js'
export { FUNCTION_NAMES } from './functions.js'
export { CONFIG_VERSION } from './config.js'
export { openSession } from './session.js'

export function compute(book: BookInput, opts: ComputeOptions): ComputeResult {
  const t0 = nowMs()
  const state = createState(book, { now: clockSerial(opts.clock), seed: opts.clock.seed, externals: opts.externals })
  return buildResult(state, opts.clock, t0)
}

export function validateFormula(text: string, ctx?: ValidateContext): ValidateResult {
  const r = validate(text, ctx)
  if (!r.ok) return r
  const { ast: _ast, ...rest } = r
  return rest
}

export function validateName(name: string): NameCheck {
  return checkName(name)
}

/** Generated code of the book in evaluation order, for reading while debugging; not part of the contract. */
export function debugGeneratedCode(book: BookInput, opts: ComputeOptions): string {
  return debugSource(createState(book, { now: clockSerial(opts.clock), seed: opts.clock.seed, externals: opts.externals }))
}

/** The contract implementation as one object. */
export const engine: FormulaEngine = { compute, validateFormula, validateName, openSession }
