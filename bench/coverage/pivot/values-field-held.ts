/**
 * The held-out set `values-field-held`, over the rows of `stock` and the books `values-field` saved with one data field
 * left: Σ Values put back after a field came beside it, after the field beside it moved, after it went to the other
 * axis, and in a book whose saved PivotTable keeps its place with one data field (`dataOnRows`, `dataPosition`).
 */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const qty: PivotOp = { value: { field: 'Qty' } }
  const price: PivotOp = { value: { field: 'Price' } }
  const read: PivotOp = { read: true }
  const c = (id: string, ops: PivotOp[]): PivotCase => ({ id, data: 'stock', ops })
  const from = (id: string, book: string, ops: PivotOp[]): PivotCase => ({ id, from: `values-field/${book}`, ops })
  return [
    c('last-field-after-drop-back', [{ rows: ['Region'] }, qty, price, { valuesOn: 'row' }, { rows: ['Product'] }, read, { dropValue: 'Sum of Qty' }, read, qty]),
    c('kept-position-drop-back', [{ rows: ['Region', 'Product'] }, qty, price, { valuesOn: 'row', position: 2 }, { move: 'Product', to: 'hidden' }, { rows: ['Product'] }, read, { dropValue: 'Sum of Price' }, read, price]),
    c('columns-middle-field-moved', [{ rows: ['Region'] }, { columns: ['Flag', 'Code'] }, qty, price, { valuesOn: 'column', position: 2 }, read, { move: 'Code', to: 'row' }, read, { dropValue: 'Sum of Qty' }, { move: 'Code', to: 'column' }, read, qty]),
    c('rows-first-to-columns', [{ rows: ['Region'] }, { columns: ['Flag'] }, qty, price, { valuesOn: 'row', position: 1 }, read, { valuesOn: 'column' }, read, { dropValue: 'Sum of Price' }, read, price]),
    c('drop-all-field-off-back', [{ rows: ['Region', 'Product'] }, qty, price, { valuesOn: 'row', position: 1 }, read, { dropValue: 'Sum of Qty' }, { dropValue: 'Sum of Price' }, { move: 'Region', to: 'hidden' }, read, price, qty]),
    from('file-rows-last', 'rows-drop-back-2', [qty]),
    from('file-rows-first', 'rows-first-drop-back-2', [price]),
    from('file-columns-first', 'columns-first-drop-back-2', [price]),
  ]
}
