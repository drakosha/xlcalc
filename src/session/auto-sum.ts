/**
 * Excel's AutoSum at a cell (`Session.getAutoSum`): the numbers right above the cell, else right beside it on the
 * left, by the values the engine holds, and the formula Excel enters over them, `SUM` of that range; `SUM()` without
 * numbers next to the cell.
 */
import type { AutoSumGuess } from '../contract.js'
import type { EngineState } from '../engine.js'
import { showFormula } from '../locale/index.js'
import { DEFAULT_LOCALE, LOCALES } from '../locale/registry.js'
import { formatAddress } from '../model.js'
import { cellValue } from '../result.js'

export function autoSumOf(state: EngineState, s: number, r: number, c: number): AutoSumGuess {
  const isNumber = (row: number, col: number) => row >= 0 && col >= 0 && typeof cellValue(state, s, row, col) === 'number'
  let range: string | null = null
  let k = r - 1
  while (isNumber(k, c)) k--
  if (k < r - 1) range = span(k + 1, c, r - 1, c)
  else {
    let j = c - 1
    while (isNumber(r, j)) j--
    if (j < c - 1) range = span(r, j + 1, r, c - 1)
  }
  const formula = `SUM(${range ?? ''})`
  const shown = showFormula(formula, LOCALES.get(state.input.locale ?? DEFAULT_LOCALE)!, undefined, state.input.externalBooks)
  return { range, formula, formulaLocal: shown.ok ? shown.formula : formula }
}

/** A range's address, a cell's alone when it is one. */
const span = (r1: number, c1: number, r2: number, c2: number) => (r1 === r2 && c1 === c2 ? formatAddress(r1, c1) : `${formatAddress(r1, c1)}:${formatAddress(r2, c2)}`)
