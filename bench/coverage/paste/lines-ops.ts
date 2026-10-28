/**
 * Paste Special's operations and Skip Blanks over whole columns and rows, onto a table and over rows a filter hid,
 * predicted by the engine before Excel ran them (`lines-ops.predicted.json`): a column and two added onto, subtracted
 * from and skipped over columns of numbers, a formula, text and blanks; a row multiplied into a row; numbers added onto
 * a table's body, its header row and its calculated column, a table's column copied out with an operation and Skip
 * Blanks; on a filtered sheet, numbers added over hidden rows, a source over hidden rows added and with Skip Blanks, a
 * whole column added. The targets' formulas go to Excel as a typed formula is saved (`plain`), as Excel refuses an
 * operation onto a dynamic array formula.
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import type { PasteCase, PasteCases, PasteOperation, PasteWhat } from './cases.js'

/** S: a source A1:A6 (A3 blank, A4 text, A5 a formula) and B1:B6; targets C1:C6 and E1:F6 of numbers, a formula, text and blanks; rows 8 and 9 of numbers. */
function linesBook(): Record<string, CellInput> {
  const cells: Record<string, CellInput> = {
    A1: { value: 2 },
    A2: { value: 3, style: 1 },
    A4: { value: 'a' },
    A5: { formula: 'B1*2' },
    A6: { value: 0.5 },
    B1: { value: 4 },
    B2: { value: -1 },
    B6: { value: 1 },
    C1: { value: 10 },
    C2: { formula: 'B2+1' },
    C3: { value: 10 },
    C4: { value: 10 },
    C6: { value: 'c' },
    E1: { value: 1 },
    E2: { value: 1 },
    F1: { value: 1 },
    F3: { formula: 'E1*3' },
    F6: { value: 2 },
  }
  for (const c of 'ABCDEF') {
    cells[`${c}8`] = { value: 'ABCDEF'.indexOf(c) + 1 }
    cells[`${c}9`] = { value: 10 }
  }
  delete cells.D8
  return cells
}

/** The keys of rows 2:8 a filter reads: `x` kept, `y` hidden. */
const KEYS = 'xyxyxxx'

/** F: A1:A8 a header and numbers (A4 blank), B1:B8 the keys, C1:C8 numbers to add onto, D2 a formula; rows 3 and 5 hidden by the filter. */
function filteredCells(): Record<string, CellInput> {
  const cells: Record<string, CellInput> = { A1: { value: 'n' }, B1: { value: 'k' }, C1: { value: 'c' }, D2: { formula: 'C2*2' } }
  for (let r = 2; r <= 8; r++) {
    if (r !== 4) cells[`A${r}`] = { value: r }
    cells[`B${r}`] = { value: KEYS[r - 2] }
    cells[`C${r}`] = { value: 100 }
  }
  return cells
}

/** T: a table T over A1:C5 (x, y, z = `T[@x]*2`), numbers E1:E4 (E3 blank). */
function tableCells(): Record<string, CellInput> {
  const cells: Record<string, CellInput> = { A1: { value: 'x' }, B1: { value: 'y' }, C1: { value: 'z' }, E1: { value: 5 }, E2: { value: 6 }, E4: { value: 7 } }
  for (let r = 2; r <= 5; r++) Object.assign(cells, { [`A${r}`]: { value: r }, [`B${r}`]: { value: 10 * r }, [`C${r}`]: { formula: 'T[@x]*2' } })
  delete cells.B4
  return cells
}

export function opsBook(): BookInput {
  return {
    styles: [{}, { font: { bold: true } }],
    sheets: [
      { name: 'S', cells: linesBook() },
      { name: 'T', cells: tableCells(), tables: [{ name: 'T', range: 'A1:C5', columns: [{ name: 'x' }, { name: 'y' }, { name: 'z', formula: 'T[@x]*2' }] }] },
      { name: 'F', cells: filteredCells(), rows: { '3': { hidden: true }, '5': { hidden: true } }, filtered: true },
    ],
  }
}

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, sheet: string, from: string, to: string, what: PasteWhat, o: { operation?: PasteOperation; skipBlanks?: boolean }, toSheet = sheet) =>
    cases.push({ id, book: 'ops', from: { sheet, range: from }, to: { sheet: toSheet, range: to }, what, ...o })
  add('column-add-values', 'S', 'A:A', 'C:C', 'values', { operation: 'add' })
  add('column-add-all', 'S', 'A:A', 'C:C', 'all', { operation: 'add' })
  add('columns-subtract-formulas', 'S', 'A:B', 'E:F', 'formulas', { operation: 'subtract' })
  add('column-skip-all', 'S', 'A:A', 'C:C', 'all', { skipBlanks: true })
  add('column-skip-add', 'S', 'A:A', 'C:C', 'values', { skipBlanks: true, operation: 'add' })
  add('row-multiply-values', 'S', '8:8', '9:9', 'values', { operation: 'multiply' })
  add('row-skip-all', 'S', '8:8', '9:9', 'all', { skipBlanks: true })
  add('table-body-add', 'T', 'E1:E3', 'B2', 'values', { operation: 'add' })
  add('table-header-add', 'T', 'E1:E2', 'A1', 'all', { operation: 'add' })
  add('table-calculated-add', 'T', 'E1:E4', 'C2', 'formulas', { operation: 'add' })
  add('table-skip-all', 'T', 'E1:E4', 'B2', 'all', { skipBlanks: true })
  add('table-column-out-add', 'T', 'B2:B5', 'E1', 'values', { operation: 'add' })
  add('filtered-add-over-hidden', 'F', 'A7:A8', 'C2:C7', 'values', { operation: 'add' })
  add('filtered-source-add', 'F', 'A2:A8', 'E2', 'values', { operation: 'add' }, 'S')
  add('filtered-source-skip', 'F', 'A2:A8', 'C2', 'all', { skipBlanks: true }, 'S')
  add('filtered-column-add', 'F', 'A:A', 'C:C', 'values', { operation: 'add' })
  add('filtered-skip-over-hidden', 'F', 'A2:A4', 'C2:C7', 'all', { skipBlanks: true })
  return { books: { ops: opsBook() }, cases, plain: { ops: ['A5', 'C2', 'C3', 'C4', 'C5', 'D2', 'F3'] } }
}
