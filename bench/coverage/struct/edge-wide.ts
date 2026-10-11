/**
 * Rows and columns deleted and inserted under a column span from C to XFD (15 wide, filled): the span keeps
 * reaching XFD. Excel takes such a span's width as the sheet's standard width when it loads the book, so widths
 * are not compared; the look of every cell is.
 */
import type { StructCase, StructCases } from './cases.js'
import { STYLES } from './edge-layout.js'
import type { StoredBook, StoredCell } from '../stored.js'

/** S: columns C to XFD 15 wide, filled, with a format; rows 3 and 4 bold, 25 points high, with a format. */
function wideBook(): StoredBook {
  return {
    styles: STYLES,
    sheets: [{ name: 'S', cells: { A1: { value: 1 }, B1: { value: 2 }, C1: { value: 3 }, A3: { formula: 'SUM(A1:C1)' } }, columns: { 'C:XFD': { width: 15, style: 1, format: '0.0' } }, rows: { '3:4': { height: 25, style: 0, format: '#,##0' } } }],
  }
}

export function buildStructCases(): StructCases {
  const extra = [{ sheet: 'S', cell: 'XFD5' }]
  const cases: StructCase[] = [
    { id: 'delete-column-A', book: 'wide', op: { op: 'deleteColumns', sheet: 'S', column: 'A', count: 1 }, extra },
    { id: 'delete-column-D', book: 'wide', op: { op: 'deleteColumns', sheet: 'S', column: 'D', count: 1 }, extra },
    { id: 'insert-column-B', book: 'wide', op: { op: 'insertColumns', sheet: 'S', column: 'B', count: 1 }, extra },
    { id: 'delete-row-3', book: 'wide', op: { op: 'deleteRows', sheet: 'S', row: 3, count: 1 }, extra },
  ]
  return { books: { wide: wideBook() }, cases, look: ['A1048566:E1048576', 'XEX1:XFD7'], widths: false }
}
