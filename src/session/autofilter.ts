/**
 * The range an AutoFilter acts on (`Session.getAutoFilterRange`), as Excel reads it (`AutoFilter.Range`, the filter
 * sets `grow` and `grow-held` in both Excels): its range down to the last row of its current region, the block of
 * cells that hold something around it, reached from cell to cell, diagonally too, up to an empty row and column all
 * round, as Excel's `CurrentRegion`; its header row and columns stay. The AutoFilter's own range stays, as Excel saves it.
 */
import type { EngineState } from '../engine.js'
import { MAX_COLS, MAX_ROWS } from '../model.js'
import { parseRange, rectText } from '../structure/area.js'

/** The range sheet `s`'s AutoFilter acts on; null without one. It reads the lines around the region as it grows, a column by the chunks of its rows that hold values. */
export function autoFilterRange(state: EngineState, s: number): string | null {
  const f = state.input.sheets[s].autoFilter
  if (!f) return null
  const x = parseRange(f.range)!
  const data = state.ctx.sheets[s]
  // whether row `r` holds something in columns `c1` to `c2`, or column `c` in rows `r1` to `r2`; a spill holds the cells past its stored ones too
  const inRow = (r: number, c1: number, c2: number) => {
    if (r < 0 || r >= MAX_ROWS) return false
    for (let c = Math.max(0, c1); c <= Math.min(MAX_COLS - 1, c2); c++) if (data.cols[c]?.get(r) != null) return true
    return data.fills.some((y) => y.r1 <= r && r <= y.r2 && y.c1 <= c2 && y.c2 >= c1)
  }
  const inColumn = (c: number, r1: number, r2: number) => {
    if (c < 0 || c >= MAX_COLS) return false
    if (data.cols[c]?.holds(Math.max(0, r1), Math.min(MAX_ROWS - 1, r2))) return true
    return data.fills.some((y) => y.c1 <= c && c <= y.c2 && y.r1 <= r2 && y.r2 >= r1)
  }
  const g = { ...x }
  for (let grew = true; grew; ) {
    grew = false
    while (inRow(g.r2 + 1, g.c1 - 1, g.c2 + 1)) (g.r2++, (grew = true))
    while (inRow(g.r1 - 1, g.c1 - 1, g.c2 + 1)) (g.r1--, (grew = true))
    while (inColumn(g.c2 + 1, g.r1 - 1, g.r2 + 1)) (g.c2++, (grew = true))
    while (inColumn(g.c1 - 1, g.r1 - 1, g.r2 + 1)) (g.c1--, (grew = true))
  }
  return g.r2 > x.r2 ? rectText({ ...x, r2: g.r2 }) : f.range
}
