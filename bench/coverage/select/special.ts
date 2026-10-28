/**
 * Go To Special (`Range.SpecialCells`), predicted by the engine before Excel ran it (`special.predicted.json`): blanks,
 * constants and formulas of each kind of value and of several over a range of one cell (the used range, which a
 * formatted empty cell and a spill reach, and which starts past A1) and over larger ranges, and visible cells over
 * ranges (over one cell, the whole sheet's, Excel took past two minutes), entire
 * columns among them; a formula answering empty text, a spill's cells, an error as a formula, hidden rows and columns.
 */
import type { SpecialCells } from '../../../src/contract.js'
import type { SelectBook, SelectCase } from './cases.js'

/** Mixed contents in B2:F8, a spill from H2, a formatted empty cell at J10; row 5 and column E hidden. */
const mixed: SelectBook = {
  styles: [{ font: { bold: true } }],
  cells: {
    B2: { value: 1 }, C2: { value: 'a' }, D2: { value: true }, E2: { formula: 'NA()' }, F2: { formula: '#DIV/0!' },
    B3: { formula: 'B2*2' }, C3: { formula: '"x"&C2' }, D3: { formula: 'D2' }, E3: { formula: '1/0' }, F3: { formula: '""' },
    B4: { value: 2.5 }, D4: { value: 'b' },
    B6: { formula: 'IF(B2>0,"","n")' }, C6: { value: 7, style: 0 }, F6: { value: false },
    B8: { value: 3 }, F8: { formula: 'B8+B4' },
    H2: { formula: 'SEQUENCE(3)' },
    J10: { style: 0 },
  },
  rows: { '5': { hidden: true } },
  columns: { E: { hidden: true } },
}

/** The book four times over, four cases a book: Excel takes seconds over a case it finds no cells for, past a job's time for one book. */
export const BOOKS: Record<string, SelectBook> = { mixed, 'mixed-2': mixed, 'mixed-3': mixed, 'mixed-4': mixed }

const S = (special: SpecialCells, range: string, id: string): Omit<SelectCase, 'book'> & { special: SpecialCells } => ({ id, range, special })

export function buildSelectCases(): SelectCase[] {
  const books = Object.keys(BOOKS)
  return [
    S({ type: 'blanks' }, 'A1', 'blanks-used'),
    S({ type: 'blanks' }, 'B2:F8', 'blanks-range'),
    S({ type: 'blanks' }, 'C:C', 'blanks-column'),
    S({ type: 'constants' }, 'A1', 'constants'),
    S({ type: 'constants', values: ['numbers'] }, 'A1', 'constants-numbers'),
    S({ type: 'constants', values: ['text', 'logicals'] }, 'A1', 'constants-text-logicals'),
    S({ type: 'constants', values: ['errors'] }, 'A1', 'constants-errors'),
    S({ type: 'formulas' }, 'A1', 'formulas'),
    S({ type: 'formulas', values: ['numbers'] }, 'A1', 'formulas-numbers'),
    S({ type: 'formulas', values: ['text'] }, 'A1', 'formulas-text'),
    S({ type: 'formulas', values: ['errors', 'logicals'] }, 'B2:F8', 'formulas-errors-logicals'),
    S({ type: 'formulas', values: ['errors'] }, 'B4:D4', 'formulas-none'),
    S({ type: 'visible' }, 'A1:K12', 'visible-wide'),
    S({ type: 'visible' }, 'D4:F6', 'visible-range'),
    S({ type: 'constants' }, 'H2:H4', 'constants-spill'),
    S({ type: 'blanks' }, 'H2:H5', 'blanks-spill'),
  ].map((c, i) => ({ ...c, book: books[i % books.length] }))
}
