/**
 * Deleting rows and columns at the sheet's end under a styled layout that reaches it: rows to 1,048,576 with a
 * height and a fill, columns to XFD with a width and bold, merged areas and styled cells on the last row and
 * column. The look of the sheet's last rows and columns is read too.
 */
import type { StructCase, StructCases } from './cases.js'
import type { StoredBook, StoredCell } from '../stored.js'

export const STYLES = [{ font: { bold: true } }, { fill: { color: '#FFCC00' } }, { fill: { color: '#99CCFF' } }]

/** S: rows 1,048,570 to the end filled, 30 points high, with a format; columns XFA to XFD bold, 20 wide, with a format; merged areas at both ends. */
function endBook(): StoredBook {
  const cells: Record<string, StoredCell> = {
    A1: { value: 1 },
    A2: { value: 2 },
    A3: { value: 3 },
    B1: { formula: 'SUM(A1:A3)' },
    A1048576: { value: 9, style: 2 },
    A1048571: { value: 7 },
    XFD5: { value: 8, style: 2 },
    XFB6: { value: 6 },
  }
  return {
    styles: STYLES,
    sheets: [
      {
        name: 'S',
        cells,
        rows: { '1048570:1048576': { height: 30, style: 1, format: '0.00' } },
        columns: { 'XFA:XFD': { width: 20, style: 0, format: '0%' } },
        merges: ['C1048575:D1048576', 'XFC1:XFD2'],
      },
    ],
  }
}

export function buildStructCases(): StructCases {
  const extra = [{ sheet: 'S', cell: 'A1048576' }, { sheet: 'S', cell: 'XFD5' }]
  const cases: StructCase[] = []
  const rows = (id: string, op: 'insertRows' | 'deleteRows', row: number, count = 1) => cases.push({ id, book: 'end', op: { op, sheet: 'S', row, count }, extra })
  const cols = (id: string, op: 'insertColumns' | 'deleteColumns', column: string, count = 1) => cases.push({ id, book: 'end', op: { op, sheet: 'S', column, count }, extra })
  rows('delete-rows-3-4', 'deleteRows', 3, 2)
  rows('delete-row-1048572', 'deleteRows', 1048572)
  rows('delete-row-1048576', 'deleteRows', 1048576)
  rows('delete-rows-1048575-2', 'deleteRows', 1048575, 2)
  rows('delete-rows-1048568-4', 'deleteRows', 1048568, 4)
  cols('delete-columns-B-C', 'deleteColumns', 'B', 2)
  cols('delete-column-XFB', 'deleteColumns', 'XFB')
  cols('delete-column-XFD', 'deleteColumns', 'XFD')
  cols('delete-columns-XEZ-3', 'deleteColumns', 'XEZ', 3)
  return { books: { end: endBook() }, cases, look: ['A1048566:E1048576', 'XEX1:XFD7'] }
}
