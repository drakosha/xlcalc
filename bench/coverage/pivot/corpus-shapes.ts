/**
 * The set `corpus-shapes`, over the rows of `shapes`: the shapes of the PivotTables a corpus of workbooks laid out
 * otherwise than Excel had saved them. Sorting by a value with ties, ascending and descending, over the products and
 * within one, with value filters before it and with several value fields on the columns; the classic layout in tabular
 * form; an item of the rows whose records all lie in a hidden item of the columns, with and without the grand total of
 * the columns; a page field of several items with one left shown.
 */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const base: PivotOp[] = [{ rows: ['Region'] }, { value: { field: 'Qty' } }]
  const two: PivotOp[] = [{ rows: ['Product', 'Region'] }, { value: { field: 'Qty' } }]
  const c = (id: string, ops: PivotOp[]): PivotCase => ({ id, data: 'shapes', ops })
  const classic: PivotOp[] = [{ table: { classic: true } }, { layout: 'tabular' }]
  return [
    c('ties-down', [...base, { sort: 'Region', order: 'descending', by: 'Sum of Qty' }]),
    c('ties-up', [...base, { sort: 'Region', order: 'ascending', by: 'Sum of Qty' }]),
    c('ties-count-down', [{ rows: ['Region'] }, { value: { field: 'Qty', fn: 'count' } }, { value: { field: 'Qty' } }, { sort: 'Region', order: 'descending', by: 'Count of Qty' }]),
    c('ties-second-value', [{ rows: ['Region'] }, { value: { field: 'Qty', fn: 'count' } }, { value: { field: 'Qty' } }, { value: { field: 'Price' } }, { sort: 'Region', order: 'descending', by: 'Sum of Qty' }]),
    c('ties-inner-down', [...two, { sort: 'Region', order: 'descending', by: 'Sum of Qty' }]),
    c('ties-inner-up', [...two, { sort: 'Region', order: 'ascending', by: 'Sum of Qty' }]),
    c('ties-both-down', [{ rows: ['Region', 'Product'] }, { value: { field: 'Qty' } }, { sort: 'Region', order: 'descending', by: 'Sum of Qty' }, { sort: 'Product', order: 'descending', by: 'Sum of Qty' }]),
    c('ties-after-value-filter', [{ rows: ['Region', 'Product'] }, { value: { field: 'Qty' } }, { sort: 'Region', order: 'descending', by: 'Sum of Qty' }, { filter: 'Region', type: 'ValueIsGreaterThan', value1: 0, by: 'Sum of Qty' }]),
    c('classic-tabular', [...classic, { rows: ['Region', 'Product'] }, { value: { field: 'Qty' } }]),
    c('classic-no-subtotals', [...classic, { rows: ['Region', 'Product'] }, { value: { field: 'Qty' } }, { subtotals: 'Region', fns: 'none' }]),
    c('classic-two-values', [...classic, { rows: ['Region'] }, { value: { field: 'Qty' } }, { value: { field: 'Price' } }]),
    c('classic-columns', [...classic, { rows: ['Region'] }, { columns: ['Product'] }, { value: { field: 'Qty' } }]),
    c('classic-page', [...classic, { pages: ['Code'] }, { rows: ['Region'] }, { value: { field: 'Qty' } }]),
    c('hidden-column-row', [{ rows: ['Region'] }, { columns: ['Product'] }, { value: { field: 'Qty' } }, { hide: 'Product', items: ['Z'] }]),
    c('hidden-column-row-no-grand', [{ rows: ['Region'] }, { columns: ['Product'] }, { value: { field: 'Qty' } }, { table: { columnGrand: false } }, { hide: 'Product', items: ['Z'] }]),
    c('hidden-row-column', [{ rows: ['Product'] }, { columns: ['Region'] }, { value: { field: 'Qty' } }, { hide: 'Product', items: ['Z'] }]),
    c('page-one-of-several', [{ pages: ['Code'] }, ...base, { page: 'Code', items: ['15', '25'] }, { hide: 'Code', items: ['25'] }]),
    c('page-one-text-of-several', [{ pages: ['Product'] }, ...base, { page: 'Product', items: ['X', 'Z'] }, { hide: 'Product', items: ['Z'] }]),
  ]
}
