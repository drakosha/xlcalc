/**
 * The set `error-order`: the order of a row field's errors, each of the seven classic ones among numbers, text, a
 * boolean and a blank (the rows of `errors`), as the PivotTable lays them out by default, sorted ascending and
 * descending by label, and as a slicer over the field lists them. Predicted before Excel ran it, the PivotTables played
 * on the engine alone: `npx tsx bench/coverage/pivot/predict.ts error-order --drive` (`error-order.predicted.json`), the slicer's items
 * added by hand as the engine orders them (`slicer/0/items`).
 */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const base: PivotOp[] = [{ rows: ['Err'] }, { value: { field: 'N' } }]
  return [
    { id: 'plain', data: 'errors', ops: base },
    { id: 'up', data: 'errors', ops: [...base, { sort: 'Err', order: 'ascending' }] },
    { id: 'down', data: 'errors', ops: [...base, { sort: 'Err', order: 'descending' }] },
    { id: 'slicer', data: 'errors', ops: [...base, { slicer: 'Err', name: 'Slicer_Err' }] },
  ]
}
