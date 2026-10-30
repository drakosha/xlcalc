/**
 * The held-out set `sort-held`, over the rows of `stock`: the sorts of `sort` on other fields and in other combinations.
 */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const two: PivotOp[] = [{ rows: ['Product', 'Region'] }, { value: { field: 'Qty' } }]
  const c = (id: string, ops: PivotOp[]): PivotCase => ({ id, data: 'stock', ops })
  return [
    c('region-down', [{ rows: ['Region'] }, { value: { field: 'Qty' } }, { sort: 'Region', order: 'descending' }]),
    c('product-up', [{ rows: ['Product'] }, { value: { field: 'Qty' } }, { sort: 'Product', order: 'ascending' }]),
    c('code-up', [{ rows: ['Code'] }, { value: { field: 'Qty' } }, { sort: 'Code', order: 'ascending' }]),
    c('when-down', [{ rows: ['When'] }, { value: { field: 'Qty' } }, { sort: 'When', order: 'descending' }]),
    c('price-down', [{ rows: ['Price'] }, { value: { field: 'Qty' } }, { sort: 'Price', order: 'descending' }]),
    c('value-average-down', [{ rows: ['Region'] }, { value: { field: 'Price', fn: 'average' } }, { sort: 'Region', order: 'descending', by: 'Average of Price' }]),
    c('value-count-up-ties', [{ rows: ['Product'] }, { value: { field: 'Flag', fn: 'count' } }, { sort: 'Product', order: 'ascending', by: 'Count of Flag' }]),
    c('second-value', [{ rows: ['Region'] }, { value: { field: 'Qty' } }, { value: { field: 'Price', fn: 'max' } }, { sort: 'Region', order: 'descending', by: 'Max of Price' }]),
    c('outer-label-inner-value', [...two, { sort: 'Product', order: 'descending' }, { sort: 'Region', order: 'descending', by: 'Sum of Qty' }]),
    c('column-label-down', [{ rows: ['Region'] }, { columns: ['Code'] }, { value: { field: 'Qty' } }, { sort: 'Code', order: 'descending' }]),
    c('tabular-value', [...two, { layout: 'tabular' }, { sort: 'Product', order: 'ascending', by: 'Sum of Qty' }]),
    c('manual-two', [{ rows: ['Region'] }, { value: { field: 'Qty' } }, { item: 'North', of: 'Region', position: 1 }, { item: 'Central', of: 'Region', position: 6 }]),
    c('manual-inner', [...two, { item: 'South', of: 'Region', position: 1 }]),
    c('manual-then-sort', [{ rows: ['Region'] }, { value: { field: 'Qty' } }, { item: 'South', of: 'Region', position: 1 }, { read: true }, { sort: 'Region', order: 'ascending' }]),
    c('days-list', [...['Tue', 'Monday', 'sun', 'Sat', 'Wednesday', 'Thu', 'fri', 'Jun', 'May'].map((m, i): PivotOp => ({ edit: `D!I${2 + i}`, to: m })), { refresh: true }, { rows: ['Note'] }, { value: { field: 'Qty' } }]),
    c('refresh-value-up-new', [{ rows: ['Region'] }, { value: { field: 'Qty' } }, { sort: 'Region', order: 'ascending', by: 'Sum of Qty' }, { read: true }, { edit: 'D!A2', to: 'Middle' }, { edit: 'D!F3', to: 90 }, { refresh: true }]),
    c('refresh-label-down-gone', [{ rows: ['Region'] }, { value: { field: 'Qty' } }, { sort: 'Region', order: 'descending' }, { read: true }, ...[2, 7, 12, 17, 22, 32, 37, 42, 47, 52, 57].map((r): PivotOp => ({ edit: `D!A${r}`, to: 'North' })), { refresh: true }]),
    c('refresh-manual-gone', [{ rows: ['Region'] }, { value: { field: 'Qty' } }, { item: 'Central', of: 'Region', position: 3 }, { read: true }, { edit: 'D!A2', to: 'Zeta' }, { refresh: true }]),
  ]
}
