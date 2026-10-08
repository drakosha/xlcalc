/**
 * Whole tables copied and moved, and what Excel's AutoCorrect does as cells are typed next to and into a table (it
 * extends the table, it fills a calculated column), which a session does not do.
 */
import { book, on, sales, table, type TableCase } from './cases.js'

export function buildTableCases(): TableCase[] {
  const plain = () => table('Sales', 'A1', ['Region', 'Qty', 'Price', 'Total'], [['North', 2, 10, ''], ['South', 3, 20, ''], ['East', 1, 40, '']])
  return [
    { id: 'copy-whole', about: 'a whole table copied on its sheet: a new table, its formulas its own', book: book(['S', { J1: '=SUM(Sales[Qty])' }, sales()]), steps: [{ op: { op: 'copyRange', from: { sheet: 'S', range: 'A1:D5' }, to: { sheet: 'S', range: 'A10' } } }], read: on('S', 'A1:J14') },
    { id: 'copy-whole-other', about: 'a whole table copied to another sheet', book: book(['S', {}, sales()], ['T', {}]), steps: [{ op: { op: 'copyRange', from: { sheet: 'S', range: 'A1:D5' }, to: { sheet: 'T', range: 'B2' } } }], read: on('T', 'A1:F7') },
    { id: 'move-whole-other', about: 'a whole table cut and pasted on another sheet: the table goes there', book: book(['S', { J1: '=SUM(Sales[Total])' }, sales()], ['T', { A1: '=ROWS(Sales)' }]), steps: [{ op: { op: 'moveRange', from: { sheet: 'S', range: 'A1:D5' }, to: { sheet: 'T', cell: 'C3' } } }], read: [...on('S', 'A1:J5'), ...on('T', 'A1:F8')] },
    { id: 'type-below', about: "a value typed right below a table: Excel's AutoCorrect extends the table", book: book(['S', { J1: '=ROWS(Sales)' }, sales()]), steps: [{ edit: { sheet: 'S', cell: 'A6', value: 'West' } }], read: on('S', 'A1:J6') },
    { id: 'type-calc', about: "a formula typed into an empty column of a table: Excel's AutoCorrect fills the column", book: book(['S', {}, plain()]), steps: [{ edit: { sheet: 'S', cell: 'D2', value: { formula: '[@Qty]*[@Price]' } } }], read: on('S', 'A1:D4') },
  ]
}
