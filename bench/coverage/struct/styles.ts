/**
 * Styles under structure operations: cell styles and number formats, row heights and styles, column widths
 * and styles, merged areas. Inserting and deleting rows and columns inside, at the edges and next to them,
 * inserting under a row with styles and formats, moving and copying ranges, whole rows and columns with
 * their styles and merged areas; last, the moves and copies that cut a merged area (Excel's refusals).
 */
import { rangeAt, type StructCase, type StructCases } from './cases.js'
import type { StoredBook, StoredCell } from '../stored.js'

const f = (formula: string): StoredCell => ({ formula })

/** Bold, orange, bold on green, blue. */
export const STYLES: StoredBook['styles'] = [{ font: { bold: true } }, { fill: { color: '#FFC000' } }, { font: { bold: true }, fill: { color: '#92D050' } }, { fill: { color: '#00B0F0' } }]

/**
 * S: values in A1:A12 and B1:B12, styled cells in rows 3 and 4 (B3 bold, C3 orange, D3 `0.00`, B4:D4 bold on
 * green with C4 `0%`, D5 blue and empty), row 6 blue at height 30 with format `0`, row 8 at height 24, row 9
 * orange with format `0.000`; column C 20 wide, E bold and 14 wide, H orange with `0.0` and 12 wide; merged
 * B10:C11, I2:J2, A14:A16; formulas in F and G. T: empty but for a formula.
 */
export function styledBook(): StoredBook {
  const s: Record<string, StoredCell> = {}
  for (let r = 1; r <= 12; r++) {
    s[`A${r}`] = { value: r }
    if (r !== 10 && r !== 11) s[`B${r}`] = { value: r * 10 }
  }
  Object.assign(s, {
    B3: { value: 30, style: 0 },
    C3: { value: 'x', style: 1 },
    D3: { value: 3.5, format: '0.00' },
    B4: { value: 40, style: 2 },
    C4: { value: 0.25, style: 2, format: '0%' },
    D4: { style: 2 },
    D5: { style: 3 },
    B10: { value: 'm' },
    I2: { value: 'i' },
    A14: { value: 'a' },
    E5: { value: 'e5' },
    H5: { value: 1.25 },
    F1: f('SUM(A1:A12)'),
    F2: f('SUM(B3:D4)'),
    G1: f('A6+B10'),
    G2: f('SUM(H:H)'),
  })
  return {
    sheets: [
      { name: 'S', cells: s, columns: { C: { width: 20 }, E: { width: 14, style: 0 }, H: { width: 12, style: 1, format: '0.0' } }, rows: { 6: { height: 30, style: 3, format: '0' }, 8: { height: 24 }, 9: { style: 1, format: '0.000' } }, merges: ['I2:J2', 'B10:C11', 'A14:A16'] },
      { name: 'T', cells: { A1: f('SUM(S!A1:A12)') } },
    ],
    styles: STYLES,
  }
}

export function buildStructCases(): StructCases {
  const cases: StructCase[] = []
  const rows = (op: 'insertRows' | 'deleteRows', row: number, count = 1) => cases.push({ id: `${op === 'insertRows' ? 'insert' : 'delete'}-rows-${row}${count > 1 ? `x${count}` : ''}`, book: 'styled', op: { op, sheet: 'S', row, count } })
  const cols = (op: 'insertColumns' | 'deleteColumns', column: string, count = 1) => cases.push({ id: `${op === 'insertColumns' ? 'insert' : 'delete'}-cols-${column}${count > 1 ? `x${count}` : ''}`, book: 'styled', op: { op, sheet: 'S', column, count } })
  const move = (from: string, to: string) => {
    const t = rangeAt(to)
    cases.push({ id: `move-${from}-${to}`, book: 'styled', op: { op: 'moveRange', from: rangeAt(from), to: { sheet: t.sheet, cell: t.range } } })
  }
  const copy = (from: string, to: string) => cases.push({ id: `copy-${from}-${to}`, book: 'styled', op: { op: 'copyRange', from: rangeAt(from), to: rangeAt(to) } })

  rows('insertRows', 5) // under the styled cells of row 4
  rows('insertRows', 7) // under row 6, blue at height 30
  rows('insertRows', 9) // under row 8 at height 24
  rows('insertRows', 10) // under row 9, orange; at the first row of B10:C11
  rows('insertRows', 11) // inside B10:C11
  rows('insertRows', 12) // under the last row of B10:C11
  rows('insertRows', 1)
  rows('insertRows', 4, 2) // under row 3: bold, orange, 0.00
  rows('insertRows', 15) // inside A14:A16
  rows('deleteRows', 4)
  rows('deleteRows', 6)
  rows('deleteRows', 10) // the first row of B10:C11
  rows('deleteRows', 10, 2) // all of B10:C11
  rows('deleteRows', 15, 2) // A14:A16 down to one cell
  rows('deleteRows', 11) // the last row of B10:C11
  cols('insertColumns', 'C') // at column C, 20 wide
  cols('insertColumns', 'D') // right of C
  cols('insertColumns', 'F') // right of E, bold
  cols('insertColumns', 'J') // inside I2:J2
  cols('insertColumns', 'I') // right of H (orange, 0.0); the first column of I2:J2
  cols('insertColumns', 'A')
  cols('deleteColumns', 'C')
  cols('deleteColumns', 'E')
  cols('deleteColumns', 'J') // I2:J2 down to one cell
  cols('deleteColumns', 'B') // B10:C11 down to one column
  move('S!B3:D4', 'S!B20')
  move('S!B10:C11', 'S!E18')
  move('S!A6', 'S!A20') // a cell of the blue row
  move('S!E5', 'S!B20') // a cell of the bold column
  move('S!B3:D5', 'T!B3')
  move('S!6:6', 'S!A20') // the blue row
  move('S!E:E', 'S!G1') // the bold column
  copy('S!B3:D4', 'S!B20')
  copy('S!B10:C11', 'S!E18')
  copy('S!B3:D3', 'S!B20:D22')
  copy('S!A6', 'S!A20')
  copy('S!E:E', 'S!G1')
  copy('S!6:6', 'S!A20')
  move('S!A19:B20', 'S!B10') // onto the whole of B10:C11
  // copies landing on part of a merged area: its top-left cell, other cells, whole columns
  copy('S!A19:A20', 'S!B10')
  copy('S!A1:A3', 'S!I1')
  copy('S!D:D', 'S!C1')
  // last: the moves and copies that cut a merged area (an error, if no window)
  move('S!B10:B11', 'S!E18')
  copy('S!B10:B11', 'S!E18')
  move('S!A19:A20', 'S!C10')
  copy('S!A19:A20', 'S!C10')
  return { books: { styled: styledBook() }, cases }
}
