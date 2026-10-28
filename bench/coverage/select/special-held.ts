/**
 * Held out after the set special fitted Go To Special: where rows and columns of their own start and end the used range
 * (a styled row above the cells and one below them, a column of a width left of them and a styled one right of them),
 * a date and empty text among constants and formulas, a formula answering an error that is not a formula's own, and
 * visible cells over an entire column with rows hidden. Predicted before Excel ran it (`special-held.predicted.json`).
 */
import type { SpecialCells } from '../../../src/contract.js'
import type { SelectBook, SelectCase } from './cases.js'

/** Cells in D5:F7; row 2 styled, row 12 styled, column B of a width, column H styled; rows 6 and 9 hidden. */
const lines: SelectBook = {
  styles: [{ font: { italic: true } }],
  cells: {
    D5: { value: 45000, numberFormat: 14 }, E5: { formula: 'D5+1' }, F5: { formula: 'T(D5)' },
    D6: { value: 'x' }, F6: { formula: 'VLOOKUP(9,D5:E5,2,FALSE)' },
    D7: { formula: 'IF(D5,"","n")' }, E7: { value: 0 },
  },
  rows: { '2': { style: 0 }, '12': { style: 0 }, '6': { hidden: true }, '9': { hidden: true } },
  columns: { B: { width: 20 }, H: { style: 0 } },
}

export const BOOKS: Record<string, SelectBook> = { lines, 'lines-2': lines, 'lines-3': lines }

const S = (special: SpecialCells, range: string, id: string): Omit<SelectCase, 'book'> & { special: SpecialCells } => ({ id, range, special })

export function buildSelectCases(): SelectCase[] {
  const books = Object.keys(BOOKS)
  return [
    S({ type: 'blanks' }, 'A1', 'blanks-used'),
    S({ type: 'constants' }, 'A1', 'constants'),
    S({ type: 'formulas', values: ['text'] }, 'A1', 'formulas-text'),
    S({ type: 'formulas', values: ['errors'] }, 'A1', 'formulas-errors'),
    S({ type: 'formulas', values: ['numbers'] }, 'D:F', 'formulas-numbers-columns'),
    S({ type: 'visible' }, 'D:D', 'visible-column'),
    S({ type: 'blanks' }, 'C4:G8', 'blanks-range'),
    S({ type: 'constants', values: ['numbers'] }, '5:7', 'constants-rows'),
  ].map((c, i) => ({ ...c, book: books[i % books.length] }))
}
