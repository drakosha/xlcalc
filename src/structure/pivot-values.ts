/**
 * Σ Values, the place of a PivotTable's value fields among the fields of its rows or columns (`'values'`), as Excel
 * keeps it while value fields are added and removed: on an axis while there are two or more, gone with fewer, and its
 * place kept in `valuesOnRows` and `valuesPosition` meanwhile, so that it comes back where it was.
 */
import type { PivotAxisField, PivotTable } from '../contract.js'

/** The fields of a patch that place Σ Values. */
export interface ValuesPlacing {
  rows?: unknown
  columns?: unknown
  valuesOnRows?: unknown
  valuesPosition?: unknown
}

const marked = (list: unknown) => Array.isArray(list) && list.includes('values')

/**
 * The PivotTable `t` (a patch `set` made to it) with Σ Values where Excel puts it: where an axis of the patch names it;
 * else at the patch's `valuesOnRows` and `valuesPosition`; else where it was, or, when it was on no axis, at the place
 * kept (the end of the columns without one). With fewer than two value fields it leaves the axes, its place kept.
 */
export function placedValues(t: PivotTable, set: ValuesPlacing = {}): PivotTable {
  const rows: PivotAxisField[] = (t.rows ?? []).filter((f) => f !== 'values')
  const columns: PivotAxisField[] = (t.columns ?? []).filter((f) => f !== 'values')
  const inRows = (t.rows ?? []).indexOf('values')
  const inColumns = (t.columns ?? []).indexOf('values')
  const now = inRows >= 0 ? { onRows: true, at: inRows } : inColumns >= 0 ? { onRows: false, at: inColumns } : { onRows: t.valuesOnRows === true, at: t.valuesPosition }
  let { onRows, at } = now
  // moved by `valuesOnRows` or `valuesPosition` alone: to the end of the other axis unless a position is given
  if (!marked(set.rows) && !marked(set.columns) && (set.valuesOnRows !== undefined || set.valuesPosition !== undefined)) {
    if (set.valuesOnRows !== undefined) onRows = t.valuesOnRows === true
    at = set.valuesPosition !== undefined ? t.valuesPosition : onRows === now.onRows ? now.at : undefined
  }
  const axis = onRows ? rows : columns
  // the last place is kept as no position, as Excel writes it
  const position = at !== undefined && at < axis.length ? at : undefined
  const out: PivotTable = { ...t }
  if ((t.values?.length ?? 0) >= 2) axis.splice(position ?? axis.length, 0, 'values')
  // an axis the PivotTable names keeps its list, an empty one too
  for (const [key, list] of [['rows', rows], ['columns', columns]] as const) if (list.length || t[key]) out[key] = list
  if (onRows) out.valuesOnRows = true
  else delete out.valuesOnRows
  if (position === undefined) delete out.valuesPosition
  else out.valuesPosition = position
  return out
}
