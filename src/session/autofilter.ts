/**
 * The range an AutoFilter acts on (`Session.getAutoFilterRange`), as Excel reads it (`AutoFilter.Range`, the filter
 * set `grow` in both Excels): its range grown down over the rows right under it that hold something in its columns or
 * the column on either side, one after another; the AutoFilter's own range stays, as Excel saves it.
 */
import type { EngineState } from '../engine.js'
import { MAX_COLS, MAX_ROWS } from '../model.js'
import { parseRange, rectText } from '../structure/area.js'

/** The range sheet `s`'s AutoFilter acts on; null without one. It reads the rows under the range up to the first that holds nothing. */
export function autoFilterRange(state: EngineState, s: number): string | null {
  const f = state.input.sheets[s].autoFilter
  if (!f) return null
  const x = parseRange(f.range)!
  const data = state.ctx.sheets[s]
  const [c1, c2] = [Math.max(0, x.c1 - 1), Math.min(MAX_COLS - 1, x.c2 + 1)]
  // a spill holds the cells past its stored ones too
  const fills = data.fills.filter((y) => y.c1 <= c2 && y.c2 >= c1)
  const held = (r: number) => {
    for (let c = c1; c <= c2; c++) if (data.cols[c]?.get(r) != null) return true
    return fills.some((y) => y.r1 <= r && r <= y.r2)
  }
  let r2 = x.r2
  while (r2 + 1 < MAX_ROWS && held(r2 + 1)) r2++
  return r2 === x.r2 ? f.range : rectText({ ...x, r2 })
}
