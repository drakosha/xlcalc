/** The computed values of a book's cells, for a paste of values (a copy from a filtered sheet). */
import type { CellValue } from '../contract.js'
import { spend } from '../exec/budget.js'
import { valueAt, type EngineState } from '../engine.js'
import { outCell } from '../result.js'
import { sheetKey } from '../syntax.js'
import type { Rect } from './area.js'

/** The non-blank values of a sheet's cells in a rectangle, spilled ones too, a large spill's unstored cells included: [row, column, value]. */
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
        spend(1)
        const v = valueAt(state, s, r, c)
        if (v !== null && v !== undefined) out.push([r, c, outCell(v)])
      }
    }
    // the unstored cells of a large spill (`spill/fill.ts`) the loop above does not reach
    for (const f of sh.fills) {
      if (f.v === null) continue
      const v = outCell(f.v)
      for (let c = Math.max(x.c1, f.c1); c <= Math.min(x.c2, f.c2); c++)
        for (let r = Math.max(x.r1, f.r1); r <= Math.min(x.r2, f.r2); r++) {
          if (sh.cols[c] && r < sh.rows) continue
          spend(1)
          out.push([r, c, v])
        }
    }
    return out
  }
}
