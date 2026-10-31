/**
 * The set `source-deleted`: the sheet of a pivot cache's source deleted (a range, the table `Sales`, the name
 * `SalesData`), the PivotTable read, refreshed and read again: what Excel keeps of the source in the book it saves, which
 * `writeXlsx` writes after such a deletion.
 */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const ops: PivotOp[] = [{ rows: ['Region'] }, { value: { field: 'Qty' } }, { deleteSheet: 'D' }, { read: true }, { refresh: true }]
  return [
    { id: 'range', ops },
    { id: 'table', source: 'table', ops },
    { id: 'name', source: 'name', ops },
  ]
}
