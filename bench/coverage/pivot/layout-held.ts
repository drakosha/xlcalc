/**
 * The held-out set `layout-held`, over the rows of `stock`: the forms of `layout` in other combinations (fields on both
 * axes, two data fields, subtotals of other functions, blank lines and captions together, repeated labels with
 * subtotals at the top, grand totals off with Σ Values on the rows).
 */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const two: PivotOp[] = [{ rows: ['Product', 'Region'] }, { value: { field: 'Qty' } }]
  const three: PivotOp[] = [{ rows: ['Flag', 'Region', 'Code'] }, { value: { field: 'Price' } }]
  const c = (id: string, ops: PivotOp[]): PivotCase => ({ id, data: 'stock', ops })
  return [
    c('two-outline-columns', [...two, { columns: ['Flag'] }, { layout: 'outline' }]),
    c('two-tabular-columns', [...two, { columns: ['Flag'] }, { layout: 'tabular' }]),
    c('three-compact-two-values', [...three, { value: { field: 'Qty', fn: 'max' } }]),
    c('three-tabular-repeat-top', [...three, { layout: 'tabular' }, { repeatAll: true }, { field: 'Flag', props: { subtotalsTop: true } }]),
    c('three-outline-repeat-top', [...three, { layout: 'outline' }, { repeatAll: true }, { field: 'Flag', props: { subtotalsTop: true } }, { field: 'Region', props: { subtotalsTop: true } }]),
    c('subtotals-var', [...two, { layout: 'outline' }, { subtotals: 'Product', fns: ['var', 'stdDevp', 'countNums'] }]),
    c('subtotals-product', [...three, { subtotals: 'Region', fns: ['product', 'min'] }, { layout: 'tabular' }]),
    c('subtotals-none-blank', [...three, { layout: 'outline' }, { subtotals: 'Flag', fns: 'none' }, { field: 'Flag', props: { blankLine: true } }]),
    c('no-grand-values-rows', [...two, { value: { field: 'Price', fn: 'average' } }, { valuesOn: 'row' }, { table: { rowGrand: false, columnGrand: false } }]),
    c('no-column-grand-two-columns', [{ rows: ['Region'] }, { columns: ['Flag', 'Product'] }, { value: { field: 'Qty' } }, { table: { columnGrand: false } }]),
    c('captions-outline', [...two, { layout: 'outline' }, { table: { grandTotalName: 'Total of all' } }, { field: 'Product', props: { caption: 'Goods' } }, { field: 'Region', props: { caption: 'Place' } }]),
    c('captions-compact-columns', [...two, { columns: ['Code'] }, { table: { rowHeader: 'Lines', columnHeader: 'Codes', grandTotalName: 'Sum' } }]),
    c('blank-outline-columns', [...two, { columns: ['Flag'] }, { layout: 'outline' }, { field: 'Product', props: { blankLine: true } }]),
    c('form-mixed-repeat', [...three, { field: 'Flag', props: { form: 'outline', repeatLabels: true } }, { field: 'Region', props: { form: 'tabular' } }]),
    c('values-rows-first-tabular', [...two, { value: { field: 'Price' } }, { valuesOn: 'row', position: 1 }, { layout: 'tabular' }]),
    c('values-columns-first', [{ rows: ['Region'] }, { columns: ['Flag'] }, { value: { field: 'Qty' } }, { value: { field: 'Price', fn: 'min' } }, { valuesOn: 'column', position: 1 }]),
    c('pages-three-outline', [{ pages: ['Flag', 'Code', 'Note'] }, ...two, { layout: 'outline' }]),
    c('empty-and-error-text', [{ rows: ['Region'] }, { columns: ['Product'] }, { value: { field: 'Amount' } }, { layout: 'tabular' }, { table: { emptyText: 'none', errorText: '?' } }]),
    c('phases-forms', [...three, { read: true }, { layout: 'outline' }, { read: true }, { layout: 'tabular' }, { read: true }, { layout: 'compact' }]),
    c('phases-values-axis', [...two, { value: { field: 'Price' } }, { read: true }, { valuesOn: 'row' }, { read: true }, { move: 'Region', to: 'column' }]),
  ]
}
