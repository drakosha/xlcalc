/** The computed values of a book's cells, for a paste of values (a copy from a filtered sheet). */
import type { CellValue } from '../contract.js'
import { toCellValue, valueAt, type EngineState } from '../engine.js'
import { outValue } from '../result.js'
import { sheetKey } from '../syntax.js'
import type { Rect } from './area.js'

/** The non-blank values of a sheet's cells in a rectangle, spilled ones too: [row, column, value]. */
export type ValuesIn = (sheet: string, rect: Rect) => [number, number, CellValue][]

/** The values of a computed state. */
export function stateValues(state: EngineState): ValuesIn {
  return (sheet, x) => {
    const s = state.sheetIdx.get(sheetKey(sheet))!
    const sh = state.ctx.sheets[s]
    const out: [number, number, CellValue][] = []
    for (let c = x.c1; c <= Math.min(x.c2, sh.cols.length - 1); c++) {
      if (!sh.cols[c]) continue
      for (let r = x.r1; r <= Math.min(x.r2, sh.rows - 1); r++) {
        const v = valueAt(state, s, r, c)
        if (v !== null && v !== undefined) out.push([r, c, outValue(toCellValue(v))])
      }
    }
    return out
  }
}
