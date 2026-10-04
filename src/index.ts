/**
 * `@app/formula`: движок формул книги. Контракт в `./contract`; реализация (вариант C замера WB1)
 * компилирует формулы в JS и считает по графу ссылок.
 */
import type { BookInput, ComputeOptions, ComputeResult, NameCheck, ValidateContext, ValidateResult } from './contract'
import { createState, debugSource } from './engine'
import { buildResult, clockSerial } from './result'
import { validate, validateName as checkName } from './syntax'

export * from './contract'
export { FUNCTION_NAMES } from './functions'
export { CONFIG_VERSION } from './config'

export function compute(book: BookInput, opts: ComputeOptions): ComputeResult {
  const t0 = performance.now()
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

/** Сгенерированный код книги в порядке вычисления: для чтения глазами при отладке, не контракт. */
export function debugGeneratedCode(book: BookInput, opts: ComputeOptions): string {
  return debugSource(createState(book, { now: clockSerial(opts.clock), seed: opts.clock.seed, externals: opts.externals }))
}
