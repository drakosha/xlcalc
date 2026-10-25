/**
 * Hidden rows of a sheet as SUBTOTAL and AGGREGATE see them: spans of rows, and whether the sheet is filtered (a filter
 * has criteria), where every hidden row counts as one a filter hid. Hidden columns count as visible.
 */
import type { ColumnInput, RowInput, SheetInput } from '../contract.js'
import type { Axis } from '../structure/area.js'
import { lines } from './spans.js'

/** A row or column hidden as Excel takes it: marked hidden, or of size 0 (a row's its height, a column's its width). */
export const isHidden = (e: { hidden?: boolean; height?: number; width?: number }): boolean => e.hidden === true || e.height === 0 || e.width === 0
/** A row (`isHidden`, a row has no width). */
export const isHiddenRow = (e: RowInput): boolean => isHidden(e)
/** A column (`isHidden`, a column has no height). */
export const isHiddenColumn = (e: ColumnInput): boolean => isHidden(e)

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

const SPANS = new WeakMap<object, [number, number][]>()

/**
 * A sheet's hidden rows (`r`) or columns (`c`) as spans (`hiddenSpans`), read once for each map: a change of the layout,
 * a structure operation and an undo give the sheet a new map, never change one (`layout/lines.ts`).
 */
export function hiddenSpansOf(map: Record<string, RowInput | ColumnInput> | undefined, axis: Axis): [number, number][] {
  if (!map) return []
  let spans = SPANS.get(map)
  if (!spans) SPANS.set(map, (spans = axis === 'r' ? hiddenSpans(map, 'r', isHiddenRow) : hiddenSpans(map, 'c', isHiddenColumn)))
  return spans
}

/** A sheet's hidden rows, joined where they touch; undefined when none is hidden. */
export function hiddenRows(sheet: SheetInput): HiddenRows | undefined {
  const spans = hiddenSpans(sheet.rows, 'r', isHiddenRow)
  return spans.length ? { spans, filtered: sheet.filtered === true } : undefined
}
