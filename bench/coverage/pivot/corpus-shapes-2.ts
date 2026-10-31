/**
 * The set `corpus-shapes-2`, over the rows of `blanks`: an item of the rows whose records all lie in a hidden item of
 * another field, the empty item of a field of text and numbers hidden on the columns (with and without the grand total
 * of the columns, compact and tabular), hidden under the item on the rows, and a page field's item hidden.
 */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const cross: PivotOp[] = [{ rows: ['Region'] }, { columns: ['Product'] }, { value: { field: 'Qty' } }]
  const c = (id: string, ops: PivotOp[]): PivotCase => ({ id, data: 'blanks', ops })
  return [
    c('blank-hidden-columns', [...cross, { hide: 'Product', items: ['(blank)'] }]),
    c('blank-hidden-columns-no-grand', [...cross, { table: { columnGrand: false } }, { hide: 'Product', items: ['(blank)'] }]),
    c('blank-hidden-columns-tabular', [...cross, { layout: 'tabular' }, { table: { columnGrand: false } }, { hide: 'Product', items: ['(blank)'] }]),
    c('text-hidden-columns', [...cross, { hide: 'Product', items: ['X'] }]),
    c('blank-hidden-inner', [{ rows: ['Region', 'Product'] }, { value: { field: 'Qty' } }, { hide: 'Product', items: ['(blank)'] }]),
    c('blank-hidden-inner-tabular', [{ rows: ['Region', 'Product'] }, { value: { field: 'Qty' } }, { layout: 'tabular' }, { hide: 'Product', items: ['(blank)'] }]),
    c('page-hides-row', [{ pages: ['Flag'] }, { rows: ['Region'] }, { value: { field: 'Qty' } }, { page: 'Flag', items: ['TRUE', 'FALSE'] }, { hide: 'Flag', items: ['FALSE'] }]),
    c('page-hides-row-columns', [{ pages: ['Flag'] }, ...cross, { page: 'Flag', items: ['TRUE', 'FALSE'] }, { hide: 'Flag', items: ['FALSE'] }]),
  ]
}
