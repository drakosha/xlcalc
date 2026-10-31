/**
 * The held-out set `areas-held`, run once the engine's prediction for it was committed after `areas` was fitted: the
 * ranges of pivot tables of other shapes (three fields on rows in each form, value fields on the rows with a column
 * field, several value fields beside column fields, column fields alone over two fields, row fields alone in tabular
 * form, report filters with their wrap and order and no grand totals, the classic layout with column fields) and the
 * books Excel saved for the sets `basic`, `filter` and `sort`, each opened as it is.
 */
import { savedBooks } from './areas.js'
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const qty: PivotOp = { value: { field: 'Qty' } }
  const price: PivotOp = { value: { field: 'Price' } }
  const amount: PivotOp = { value: { field: 'Amount' } }
  const three: PivotOp = { rows: ['Region', 'Product', 'Flag'] }
  const c = (id: string, ops: PivotOp[]): PivotCase => ({ id, data: 'stock', ops })
  return [
    c('three-compact', [three, qty]),
    c('three-outline', [three, qty, { layout: 'outline' }]),
    c('three-tabular', [three, qty, { layout: 'tabular' }]),
    c('values-rows-columns', [{ rows: ['Region'] }, { columns: ['Flag'] }, qty, price, { valuesOn: 'row' }]),
    c('values-rows-first', [{ rows: ['Region'] }, qty, price, { valuesOn: 'row', position: 1 }]),
    c('three-values-columns', [{ rows: ['Region'] }, { columns: ['Flag'] }, qty, price, amount]),
    c('values-before-columns', [{ rows: ['Region'] }, { columns: ['Flag'] }, qty, price, { valuesOn: 'column', position: 1 }]),
    c('columns-two-alone', [{ columns: ['Flag', 'Code'] }]),
    c('rows-two-tabular-alone', [{ rows: ['Region', 'Product'] }, { layout: 'tabular' }]),
    c('no-rows-three-values', [{ columns: ['Code'] }, qty, price, amount]),
    c('values-alone-three', [qty, price, amount]),
    c('pages-two-wrap-one', [{ pages: ['Product', 'Flag'] }, { rows: ['Region'] }, qty, { table: { pagesWrap: 1 } }]),
    c('pages-four-across-wrap', [{ pages: ['Product', 'Flag', 'Code', 'Note'] }, { rows: ['Region'] }, qty, { table: { pagesAcross: true, pagesWrap: 3 } }]),
    c('pages-no-grand', [{ pages: ['Code'] }, { rows: ['Region'] }, { columns: ['Flag'] }, qty, { table: { rowGrand: false, columnGrand: false } }]),
    c('page-columns-alone', [{ pages: ['Product'] }, { columns: ['Flag'] }]),
    c('classic-columns', [{ table: { classic: true } }, { layout: 'tabular' }, { rows: ['Region', 'Product'] }, { columns: ['Flag'] }, qty]),
    ...savedBooks(['basic', 'filter', 'sort'], 'saved-'),
  ]
}
