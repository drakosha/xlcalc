/**
 * Excel's AutoSum at a cell (`Session.getAutoSum`, `test/excel/ui/press-held.json`): going up from the cell, past the
 * cells right above it that hold no number, the run of numbers there up to a `SUM` or `SUBTOTAL` formula (a subtotal), to the cell;
 * else the same to the left; and the formula Excel enters over them, `SUM` of that range. When the run's nearest cell
 * is a subtotal, `SUM` of the run's subtotals alone, the nearest first (a grand total). A date or a time is no number
 * here. Without numbers, `SUM()`, which Excel opens for editing and does not take.
 */
import type { AutoSumGuess } from '../contract.js'
import type { EngineState } from '../engine.js'
import { sheetLines } from '../layout/lines.js'
import { showFormula } from '../locale/index.js'
import { DEFAULT_LOCALE, LOCALES } from '../locale/registry.js'
import { formatAddress } from '../model.js'
import { showsDate } from '../numfmt/category.js'
import { cellValue } from '../result.js'

export function autoSumOf(state: EngineState, s: number, r: number, c: number): AutoSumGuess {
  const sheet = state.input.sheets[s]
  const lines = sheetLines(sheet, state)
  const isNumber = (row: number, col: number) => {
    return typeof cellValue(state, s, row, col) === 'number' && !showsDate(lines.format(row, col), state.input.locale)
  }
  // from `from` toward the edge by `by` ([row, column]): the cells skipped, then the run of numbers; null without one
  const run = (by: [number, number]): [number, number][] | null => {
    const out: [number, number][] = []
    let [i, j] = [r + by[0], c + by[1]]
    while (i >= 0 && j >= 0 && !isNumber(i, j)) [i, j] = [i + by[0], j + by[1]]
    if (i < 0 || j < 0) return null
    while (i >= 0 && j >= 0 && isNumber(i, j)) {
      out.push([i, j])
      ;[i, j] = [i + by[0], j + by[1]]
    }
    return out
  }
  const found = run([-1, 0]) ?? run([0, -1])
  const subtotal = ([i, j]: [number, number]) => /^(SUM|SUBTOTAL)\(/i.test(sheet.cells[formatAddress(i, j)]?.formula ?? '')
  const grand = !!found && subtotal(found[0])
  const subtotals = grand ? found.filter(subtotal) : []
  // a run of numbers ends at a subtotal
  const cells = found && !grand && found.some(subtotal) ? found.slice(0, found.findIndex(subtotal)) : found
  const top = cells?.[cells.length - 1]
  const range = subtotals.length ? subtotals.map(([i, j]) => formatAddress(i, j)).join(',') : top ? span(top[0], top[1], r === top[0] ? r : r - 1, c === top[1] ? c : c - 1) : null
  const formula = `SUM(${range ?? ''})`
  const shown = showFormula(formula, LOCALES.get(state.input.locale ?? DEFAULT_LOCALE)!, undefined, state.input.externalBooks)
  return { range, formula, formulaLocal: shown.ok ? shown.formula : formula }
}

/** A range's address, a cell's alone when it is one. */
const span = (r1: number, c1: number, r2: number, c2: number) => (r1 === r2 && c1 === c2 ? formatAddress(r1, c1) : `${formatAddress(r1, c1)}:${formatAddress(r2, c2)}`)
