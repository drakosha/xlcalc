/**
 * The set `corpus-shapes-3`, over the rows of `blanks2`: the empty item of the rows whose records all hold the empty item
 * of the columns, hidden there (tabular and compact, with and without the grand total of the columns); and items
 * collapsed on the rows with a page field's item hidden and a field on the columns.
 */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const cross: PivotOp[] = [{ rows: ['Region'] }, { columns: ['Product'] }, { value: { field: 'Qty' } }]
  const c = (id: string, ops: PivotOp[]): PivotCase => ({ id, data: 'blanks2', ops })
  const hide: PivotOp = { hide: 'Product', items: ['(blank)'] }
  return [
    c('blank-row-tabular-no-grand', [...cross, { layout: 'tabular' }, { table: { columnGrand: false } }, hide]),
    c('blank-row-tabular', [...cross, { layout: 'tabular' }, hide]),
    c('blank-row-compact-no-grand', [...cross, { table: { columnGrand: false } }, hide]),
    c('blank-row-compact', [...cross, hide]),
    c('collapsed-page-hidden', [{ pages: ['Flag'] }, { rows: ['Region', 'Product'] }, { columns: ['Code'] }, { value: { field: 'Qty' } }, { page: 'Flag', items: ['TRUE', 'FALSE'] }, { hide: 'Flag', items: ['FALSE'] }, { collapse: 'Region', items: ['A'] }]),
    c('collapsed-inner-page-hidden', [{ pages: ['Flag'] }, { rows: ['Region', 'Product', 'Code'] }, { value: { field: 'Qty' } }, { page: 'Flag', items: ['TRUE', 'FALSE'] }, { hide: 'Flag', items: ['FALSE'] }, { collapse: 'Product', items: ['X'] }]),
  ]
}
