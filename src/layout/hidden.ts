/**
 * Hidden rows of a sheet as SUBTOTAL and AGGREGATE see them: spans of rows, and whether the sheet is filtered (a filter
 * has criteria), where every hidden row counts as one a filter hid. Hidden columns count as visible.
 */
import type { ColumnInput, RowInput, SheetInput } from '../contract.js'
import type { Axis } from '../structure/area.js'
import { lines } from './spans.js'

/** A row hidden as Excel takes it: marked hidden, or of height 0 (an xlsx round trip through Excel, the book `zero-sizes`). */
export const isHiddenRow = (e: RowInput): boolean => e.hidden === true || e.height === 0
/** A column hidden: marked hidden, or of width 0. */
export const isHiddenColumn = (e: ColumnInput): boolean => e.hidden === true || e.width === 0

export interface HiddenRows {
  /** Disjoint spans of hidden rows, ascending, zero-based and inclusive. */
  spans: [number, number][]
  filtered: boolean
}

/** The spans of hidden lines of a sheet's rows or columns, ascending, joined where they touch. */
export function hiddenSpans<T>(map: Record<string, T> | undefined, axis: Axis, hidden: (e: T) => boolean): [number, number][] {
  const spans: [number, number][] = []
  for (const [lo, hi, e] of lines(map, axis)) {
    if (!hidden(e)) continue
    const last = spans.at(-1)
    if (last && last[1] + 1 === lo) last[1] = hi
    else spans.push([lo, hi])
  }
  return spans
}

/** A sheet's hidden rows, joined where they touch; undefined when none is hidden. */
export function hiddenRows(sheet: SheetInput): HiddenRows | undefined {
  const spans = hiddenSpans(sheet.rows, 'r', isHiddenRow)
  return spans.length ? { spans, filtered: sheet.filtered === true } : undefined
}
