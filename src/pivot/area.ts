/** A PivotTable's cells on its sheet: its range and, above it, its Filters area. */
import type { PivotTable } from '../contract.js'
import { parseRange, type Rect } from '../structure/area.js'

/**
 * The rectangle of a PivotTable's cells: its range, and the rows of its Filters area above it with the blank row between
 * (`pageRows` + 1), each line of filters two cells wide and one apart (`pageColumns`); null for a range that does not parse.
 */
export function pivotArea(t: Pick<PivotTable, 'location'>): Rect | null {
  const x = parseRange(t.location.range)
  if (!x) return null
  const rows = t.location.pageRows ?? 0
  if (!rows) return x
  const wide = 3 * Math.max(1, t.location.pageColumns ?? 1) - 2
  return { r1: Math.max(0, x.r1 - rows - 1), c1: x.c1, r2: x.r2, c2: Math.max(x.c2, x.c1 + wide) }
}

/** The parts of a PivotTable an entry cannot change: its range and its Filters area, not the blank row between them, which takes anything as in Excel. */
export function pivotParts(t: Pick<PivotTable, 'location'>): Rect[] {
  const whole = pivotArea(t)
  const x = parseRange(t.location.range)
  if (!whole || !x) return []
  return whole.r1 < x.r1 - 1 ? [x, { ...whole, r2: x.r1 - 2 }] : [x]
}
