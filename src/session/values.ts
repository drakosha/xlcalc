/**
 * A rectangle's values (`Session.getValues`), read from the sheet's columns as rows: no result object and no address is
 * made for a cell, and each value is the one `readCells` gives in its field `value`.
 */
import type { CellValue, ValuesOptions } from '../contract.js'
import type { EngineState } from '../engine.js'
import { allotCells, CalcError, paced, type Steps } from '../exec/budget.js'
import { checkNumber, object, onlyOptions } from '../input/basic.js'
import { MAX_COLS, MAX_ROWS } from '../model.js'
import { fillAt } from '../spill/fill.js'
import { outCell } from '../result.js'
import type { Rect } from '../structure/area.js'

/** The most cells a rectangle of values may have by default: a sheet of tens of thousands of rows fits, a whole column does not. */
const MAX_CELLS = 1_000_000

/** `getValues`'s options checked. */
export function checkValuesOptions(v: unknown): ValuesOptions {
  if (v === undefined) return {}
  const o = object(v, 'opts', 'options such as { maxCells: 2000000 }')
  onlyOptions(o, 'opts', ['maxCells'])
  if (o.maxCells !== undefined) checkNumber(o.maxCells, 'opts.maxCells', 1, MAX_ROWS * MAX_COLS, true)
  return o as ValuesOptions
}

/** The values of rectangle `x` of sheet `s`, by rows; refused past `maxCells`. */
export function* valuesIn(state: EngineState, s: number, x: Rect, o: ValuesOptions): Steps<CellValue[][]> {
  const width = x.c2 - x.c1 + 1
  const size = (x.r2 - x.r1 + 1) * width
  const max = o.maxCells ?? MAX_CELLS
  if (size > max) throw new CalcError('limit', `the range has ${size} cells, over the limit of ${max}`, { limit: 'cells', max, size })
  // rows of values, or text: no one array of them all
  allotCells(size)
  const { cols, fills } = state.ctx.sheets[s]
  const out: CellValue[][] = []
  for (let r = x.r1; r <= x.r2; r++) {
    if (paced(width)) yield
    const row: CellValue[] = []
    // the cells a large spill does not store hold its fill
    for (let c = x.c1; c <= x.c2; c++) row.push(outCell(cols[c]?.get(r) ?? (fills.length ? fillAt(fills, r, c) : null)))
    out.push(row)
  }
  return out
}
