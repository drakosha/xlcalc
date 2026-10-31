/**
 * The set `slicer-file`: books of `slicer-delete` as Excel saved them after the deletion, opened and refreshed. The
 * timeline's book still holds the timeline's date filter on the PivotTable, which Excel laid out without it; the
 * slicer's holds its hidden items. Predicted before Excel ran it, in `slicer-file.predicted.json`: the cells as saved
 * when opened, and after Refresh the PivotTable filtered by what the file holds (the dates of the deleted timeline,
 * West alone).
 */
import type { PivotCase } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  return ['delete-timeline', 'delete-slicer'].map((id) => ({ id, file: `test/excel/win/pivot/slicer-delete/${id}.xlsx`, ops: [{ read: true }, { refresh: true }] }))
}
