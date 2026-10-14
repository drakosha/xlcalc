/**
 * Structure operations over hidden rows and columns: copies, fills and moves of ranges with rows hidden by hand and
 * on a filtered sheet, inserting and deleting rows inside, before and after a hidden block, inserting and deleting
 * columns next to hidden ones. Copies are applied as copy and paste; the `fill` cases record Excel's Fill Down and
 * fill handle over the same ranges.
 */
import type { BookInput, CellInput, RowInput, SheetInput } from '../../../src/contract.js'
import type { StructCase, StructCases } from './cases.js'
import { formulaCell as f } from '../stored.js'

/** The keys of rows 2:16 an autofilter reads: `x` kept, `y` hidden by the filter. */
const KEYS = 'xxxxxyyxxyxxxxx'

/** S: a header, A2:A16 the row numbers, B2:B16 the keys, C2:C16 `=A*10`; D2:E2 the top row of the fills. T: empty. */
function sheets(rows: Record<string, RowInput>, filtered = false): SheetInput[] {
  const cells: Record<string, CellInput> = { A1: { value: 'n' }, B1: { value: 'k' }, C1: { value: 'f' }, D2: { value: 'top' }, E2: f('A2+1000') }
  for (let r = 2; r <= 16; r++) Object.assign(cells, { [`A${r}`]: { value: r }, [`B${r}`]: { value: KEYS[r - 2] }, [`C${r}`]: f(`A${r}*10`) })
  return [{ name: 'S', cells, rows, ...(filtered ? { filtered: true } : {}) }, { name: 'T', cells: {} }]
}

/** Rows 7:8 and 11 (keys `y`) hidden, and row 4 (`x`). */
const HIDDEN: Record<string, RowInput> = { '4': { hidden: true }, '7:8': { hidden: true }, '11': { hidden: true } }

/** Columns C:D hidden (10 wide), values in A1:F3. */
function columnsBook(): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= 3; r++) for (const [i, c] of ['A', 'B', 'C', 'D', 'E', 'F'].entries()) cells[`${c}${r}`] = { value: r * 10 + i }
  cells.G1 = f('SUM(A1:F1)')
  return { sheets: [{ name: 'S', cells, columns: { 'C:D': { width: 10, hidden: true } } }, { name: 'T', cells: {} }] }
}

export function buildStructCases(): StructCases {
  const books: Record<string, BookInput> = {
    manual: { sheets: sheets(HIDDEN) },
    filtered: { sheets: sheets(HIDDEN, true) },
    columns: columnsBook(),
  }
  const files = { filtered: { sheets: { S: { filter: { range: 'A1:C16', keep: { 1: ['x'] } } } } } }
  const from = (range: string) => ({ sheet: 'S', range })
  const cases: StructCase[] = []
  for (const book of ['manual', 'filtered']) {
    const add = (id: string, op: StructCase['op'], via?: StructCase['via']) => cases.push({ id: `${book}-${id}`, book, op, ...(via ? { via } : {}) })
    add('copy', { op: 'copyRange', from: from('A2:C16'), to: from('G2:I16') })
    add('copy-other', { op: 'copyRange', from: from('A2:C16'), to: { sheet: 'T', range: 'A2:C16' } })
    add('copy-rows', { op: 'copyRange', from: from('2:16'), to: { sheet: 'T', range: '2:16' } })
    add('copy-onto', { op: 'copyRange', from: from('D2:E2'), to: from('D3:E16') })
    add('fill-down', { op: 'copyRange', from: from('D2:E2'), to: from('D2:E16') }, 'fillDown')
    add('fill-handle', { op: 'copyRange', from: from('D2:E2'), to: from('D2:E16') }, 'autofill')
    add('move', { op: 'moveRange', from: from('A2:C16'), to: { sheet: 'S', cell: 'G2' } })
    add('insert-before', { op: 'insertRows', sheet: 'S', row: 7 })
    add('insert-inside', { op: 'insertRows', sheet: 'S', row: 8, count: 2 })
    add('insert-after', { op: 'insertRows', sheet: 'S', row: 9 })
    add('delete-inside', { op: 'deleteRows', sheet: 'S', row: 8 })
    add('delete-around', { op: 'deleteRows', sheet: 'S', row: 6, count: 4 })
    add('copy-cell', { op: 'copyRange', from: from('A2:C16'), to: { sheet: 'T', range: 'A2' } })
    add('copy-visible', { op: 'copyRange', from: from('A2:C3'), to: { sheet: 'T', range: 'A2:C3' } })
    add('copy-over-hidden', { op: 'copyRange', from: from('A2:C3'), to: from('G6:I7') })
    add('copy-same-cell', { op: 'copyRange', from: from('A2:C16'), to: from('G2') })
    add('insert-hidden', { op: 'insertRows', sheet: 'S', row: 7, count: 2 })
    add('delete-hidden', { op: 'deleteRows', sheet: 'S', row: 7, count: 2 })
    add('insert-columns', { op: 'insertColumns', sheet: 'S', column: 'B', count: 2 })
    add('tile-over-hidden', { op: 'copyRange', from: from('A2:C3'), to: from('G2:I9') })
    add('copy-tiled', { op: 'copyRange', from: from('A2:C16'), to: { sheet: 'T', range: 'A2:C23' } })
    add('copy-below', { op: 'copyRange', from: from('A2:C16'), to: from('G12:I22') })
  }
  cases.push(
    { id: 'columns-insert-before', book: 'columns', op: { op: 'insertColumns', sheet: 'S', column: 'C' } },
    { id: 'columns-insert-inside', book: 'columns', op: { op: 'insertColumns', sheet: 'S', column: 'D' } },
    { id: 'columns-insert-after', book: 'columns', op: { op: 'insertColumns', sheet: 'S', column: 'E' } },
    { id: 'columns-delete', book: 'columns', op: { op: 'deleteColumns', sheet: 'S', column: 'C' } },
    { id: 'columns-copy', book: 'columns', op: { op: 'copyRange', from: { sheet: 'S', range: 'A1:F3' }, to: { sheet: 'S', range: 'A10:F12' } } },
  )
  return { books, cases, files }
}
