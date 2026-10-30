/** The cells a PivotTable takes on its sheet: its range and the Filters area above it, with the blank row between. */
import type { PivotLocation } from '../contract.js'
import { parseRange, type Rect } from '../structure/area.js'

/** A PivotTable's whole area (Excel's `TableRange2`): its range, and the rows of its Filters area and the row under them above it, each pair of its columns a field and its item, one column between pairs. */
export function pivotArea(location: PivotLocation): Rect | null {
  const x = parseRange(location.range)
  if (!x) return null
  const rows = location.pageRows ?? 0
  if (!rows) return x
  const cols = (location.pageColumns ?? 1) * 3 - 1
  return { r1: Math.max(0, x.r1 - rows - 1), c1: x.c1, r2: x.r2, c2: Math.max(x.c2, x.c1 + cols - 1) }
}
