/**
 * Hidden rows of a sheet as SUBTOTAL and AGGREGATE see them: spans of rows, and whether the sheet is filtered (a filter
 * has criteria), where every hidden row counts as one a filter hid. Hidden columns count as visible.
 */
import type { RowInput, SheetInput } from '../contract.js'
import { lines } from './spans.js'

/** A row hidden as Excel takes it: marked hidden, or of height 0 (an xlsx round trip through Excel, the book `zero-sizes`). */
export const isHiddenRow = (e: RowInput): boolean => e.hidden === true || e.height === 0

export interface HiddenRows {
  /** Disjoint spans of hidden rows, ascending, zero-based and inclusive. */
  spans: [number, number][]
  filtered: boolean
}

/** A sheet's hidden rows, joined where they touch; undefined when none is hidden. */
export function hiddenRows(sheet: SheetInput): HiddenRows | undefined {
  const spans: [number, number][] = []
  for (const [lo, hi, e] of lines(sheet.rows, 'r')) {
    if (!isHiddenRow(e)) continue
    const last = spans.at(-1)
    if (last && last[1] + 1 === lo) last[1] = hi
    else spans.push([lo, hi])
  }
  return spans.length ? { spans, filtered: sheet.filtered === true } : undefined
}
