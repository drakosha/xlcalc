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
 * The PivotTable `t` (a patch `set` made to it) with Σ Values where Excel puts it. Where an axis of the patch names it,
 * that is its place, kept as `valuesPosition` unless it is last; moved by the patch's `valuesOnRows` (to the end of that
 * axis) or `valuesPosition`, it goes there; else it stays where it is. On no axis, it comes at the place kept, the end of
 * the axis without one. With fewer than two value fields it leaves the axes and `valuesOnRows` and `valuesPosition` stay,
 * as Excel keeps them; a field moved beside it leaves `valuesPosition` as it was, as Excel does.
 */
export function placedValues(t: PivotTable, set: ValuesPlacing = {}): PivotTable {
  const rows: PivotAxisField[] = (t.rows ?? []).filter((f) => f !== 'values')
  const columns: PivotAxisField[] = (t.columns ?? []).filter((f) => f !== 'values')
  const inRows = (t.rows ?? []).indexOf('values')
  const inColumns = (t.columns ?? []).indexOf('values')
  const marker = inRows >= 0 || inColumns >= 0
  let onRows = marker ? inRows >= 0 : t.valuesOnRows === true
  let position = t.valuesPosition
  let at = marker ? Math.max(inRows, inColumns) : position
  if (marked(set.rows) || marked(set.columns)) position = at! < (onRows ? rows : columns).length ? at : undefined
  else if (set.valuesOnRows !== undefined || set.valuesPosition !== undefined) {
    // moved to the end of the other axis, or to the position given
    const was = onRows
    if (set.valuesOnRows !== undefined) onRows = t.valuesOnRows === true
    if (set.valuesPosition === undefined && onRows !== was) position = undefined
    if (set.valuesPosition !== undefined || onRows !== was) at = position
  }
  const axis = onRows ? rows : columns
  if ((t.values?.length ?? 0) >= 2) axis.splice(Math.min(at ?? axis.length, axis.length), 0, 'values')
  const out: PivotTable = { ...t }
  // an axis the PivotTable names keeps its list, an empty one too
  for (const [key, list] of [['rows', rows], ['columns', columns]] as const) if (list.length || t[key]) out[key] = list
  if (onRows) out.valuesOnRows = true
  else delete out.valuesOnRows
  if (position === undefined) delete out.valuesPosition
  else out.valuesPosition = position
  return out
}

/**
 * A PivotTable as a workbook gives it (`xlcalc/pivot`'s intake, `pivot/hook.ts`): one of two value fields or more with
 * Σ Values on no axis, a host's workbook or one saved before 0.72, with it where adding the second value field puts it
 * (`placedValues`); itself otherwise.
 */
export const valuesTaken = (t: PivotTable): PivotTable => ((t.values?.length ?? 0) < 2 || t.rows?.includes('values') || t.columns?.includes('values') ? t : placedValues(t))
