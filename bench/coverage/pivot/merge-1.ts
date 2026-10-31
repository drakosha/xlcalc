/**
 * The set `merge-1`: Merge and center cells with labels (`MergeLabels`) in the three forms, with two fields on the rows,
 * on the columns, and with subtotals at the bottom: the cells Excel shows and the merged areas and alignment of the book
 * it saved, which `placePivots` writes as Excel does.
 */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const base: PivotOp[] = [{ rows: ['Region', 'Product'] }, { value: { field: 'Qty' } }, { table: { mergeLabels: true } }]
  return [
    { id: 'compact', ops: base },
    { id: 'outline', ops: [...base, { layout: 'outline' }] },
    { id: 'tabular', ops: [...base, { layout: 'tabular' }] },
    { id: 'columns', ops: [{ rows: ['Region'] }, { columns: ['Flag', 'Code'] }, { value: { field: 'Qty' } }, { table: { mergeLabels: true } }] },
    { id: 'tabular-bottom', ops: [...base, { layout: 'tabular' }, { subtotals: 'Region', fns: 'automatic' }] },
  ]
}
