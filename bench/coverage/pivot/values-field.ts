/**
 * The set `values-field`: Σ Values, the place of the data fields on the rows or columns, as data fields are added and
 * taken off: where Excel puts it for the second data field, where it stays for a third, where it goes when one data
 * field is left and where it comes back with a second again, from the rows, the middle of the rows and the start of the
 * columns, and when a field beside it moves.
 */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const qty: PivotOp = { value: { field: 'Qty' } }
  const price: PivotOp = { value: { field: 'Price' } }
  const read: PivotOp = { read: true }
  const c = (id: string, ops: PivotOp[]): PivotCase => ({ id, ops })
  return [
    c('second-after-columns', [{ rows: ['Region'] }, { columns: ['Flag'] }, qty, read, price]),
    c('second-no-columns', [{ rows: ['Region'] }, qty, read, price]),
    c('third-rows-middle', [{ rows: ['Region', 'Product'] }, qty, price, { valuesOn: 'row', position: 2 }, read, { value: { field: 'Amount', fn: 'count' } }]),
    c('columns-first-drop-back', [{ rows: ['Region'] }, { columns: ['Flag'] }, qty, price, { valuesOn: 'column', position: 1 }, read, { dropValue: 'Sum of Price' }, read, price]),
    c('rows-drop-back', [{ rows: ['Region'] }, qty, price, { valuesOn: 'row' }, read, { dropValue: 'Sum of Qty' }, read, { value: { field: 'Amount', fn: 'count' } }]),
    c('rows-first-drop-back', [{ rows: ['Region', 'Product'] }, qty, price, { valuesOn: 'row', position: 1 }, read, { dropValue: 'Sum of Price' }, read, price]),
    c('rows-drop-all-back', [{ rows: ['Region'] }, qty, price, { valuesOn: 'row' }, read, { dropValue: 'Sum of Qty' }, { dropValue: 'Sum of Price' }, read, qty, price]),
    c('rows-middle-field-off', [{ rows: ['Region', 'Product'] }, qty, price, { valuesOn: 'row', position: 2 }, read, { move: 'Product', to: 'hidden' }, read, { rows: ['Product'] }]),
    c('rows-last-field-added', [{ rows: ['Region'] }, qty, price, { valuesOn: 'row' }, read, { rows: ['Product'] }]),
    c('drop-then-field-added', [{ rows: ['Region', 'Product'] }, qty, price, { valuesOn: 'row', position: 1 }, read, { dropValue: 'Sum of Qty' }, { rows: ['Flag'] }, read, qty]),
  ]
}
