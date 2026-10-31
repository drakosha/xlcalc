/**
 * The set `slicer-delete`: what a PivotTable shows after its slicer or timeline is deleted (`SlicerCache.Delete`), and
 * what becomes of a slicer or timeline when the sheet of a PivotTable it filters is deleted, the slicer on another
 * sheet: over that PivotTable alone, and over it and another. Predicted before Excel ran it: the engine's answers of the
 * deletions in `slicer-delete.predicted.json` (the PivotTable keeps the filter); the deleted sheet, which the session's
 * PivotTables do not follow yet, by the rule taken for them: a slicer or timeline goes with the last PivotTable it
 * filters (`sheet-one`, `sheet-timeline`: read, it is Excel's error), and over two it stays over the one left with its
 * choice (`sheet-two`: PT1 South alone).
 */
import { serial } from './data.js'
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const base: PivotOp[] = [{ rows: ['Region'] }, { value: { field: 'Price' } }]
  const other: PivotOp[] = [...base, { addSheet: 'Q' }, { create: 'PT2', at: 'Q!A3' }, { rows: ['Product'], pt: 'PT2' }, { value: { field: 'Qty' }, pt: 'PT2' }]
  return [
    { id: 'delete-slicer', ops: [...base, { slicer: 'Region', name: 'Slicer_Region' }, { select: 'Slicer_Region', items: ['West'] }, { read: true }, { deleteSlicer: 'Slicer_Region' }] },
    { id: 'delete-timeline', ops: [...base, { timeline: 'Date', name: 'Timeline_Date', from: serial(2024, 3, 1), to: serial(2024, 12, 31) }, { read: true }, { deleteSlicer: 'Timeline_Date' }] },
    { id: 'sheet-one', ops: [...other, { slicer: 'Region', name: 'Slicer_Region', pt: 'PT2' }, { select: 'Slicer_Region', items: ['South'] }, { read: true }, { deleteSheet: 'Q' }] },
    { id: 'sheet-two', ops: [...other, { slicer: 'Region', name: 'Slicer_Region', pivots: ['PT2'] }, { select: 'Slicer_Region', items: ['South'] }, { read: true }, { deleteSheet: 'Q' }] },
    { id: 'sheet-timeline', ops: [...other, { timeline: 'Date', name: 'Timeline_Date', pt: 'PT2', from: serial(2024, 3, 1), to: serial(2024, 12, 31) }, { read: true }, { deleteSheet: 'Q' }] },
  ]
}
