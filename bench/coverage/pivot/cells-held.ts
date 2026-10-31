/**
 * The held-out set `cells-held`, run once the engine's prediction for it was committed after `cells` was fitted: what
 * each cell is in the pivot tables Excel saved for `layout` (every form, subtotals at the top and bottom, blank lines,
 * repeated labels, captions), the built ones of `areas-held` and those of `show`.
 */
import { kindsOf } from './cells.js'
import type { PivotCase } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  return kindsOf(['layout', 'areas-held', 'show'], (id) => !id.startsWith('areas-held-saved-'))
}
