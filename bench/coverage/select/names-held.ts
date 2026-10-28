/**
 * Held out after the set names fitted Create Names from Selection: spaces at the ends and inside, a tab, R1C1
 * references of each form and in lower case, a time and a percentage, text that reads as a number, a formula's number
 * and text, FALSE, a label of spaces alone, a function's name, a range's address, ru-RU's boolean in en-US, two labels
 * of one name in a column, a label of 255 characters and one of 256. Predicted before Excel ran it
 * (`names-held.predicted.json`).
 */
import type { CellInput } from '../../../src/contract.js'
import type { SelectBook, SelectCase } from './cases.js'

const LABELS: CellInput[] = [
  { value: 'lead ' }, { value: '  two  words ' }, { value: 'tab\there' }, { value: 'R2' }, { value: 'C3' }, { value: 'RC' }, { value: 'R1C' }, { value: 'r2c3' }, { value: 'c' },
  { value: 0.5, numberFormat: 'h:mm' }, { value: 0.5, numberFormat: '0%' }, { value: '123' }, { formula: '1+1' }, { formula: '"t"&"x"' }, { value: false }, { value: '   ' },
  { value: 'SUM' }, { value: 'A1:B2' }, { value: 'ИСТИНА' }, { value: 'y'.repeat(255) }, { value: 'z'.repeat(256) }, { value: 'XFD1' }, { value: 'XFE1' },
]

const COLS = 'BCDEFGHIJKLMNOPQRSTUVWXYZ'

function labelled(): SelectBook {
  const cells: Record<string, CellInput> = { A2: { value: 'east' }, A3: { value: 'east' }, A4: { value: 'west' } }
  LABELS.forEach((x, i) => {
    cells[`${COLS[i]}1`] = x
    for (let r = 2; r <= 4; r++) cells[`${COLS[i]}${r}`] = { value: r * 10 + i }
  })
  return { cells }
}

export const BOOKS: Record<string, SelectBook> = { labels: labelled() }

export function buildSelectCases(): SelectCase[] {
  return [
    { id: 'top', book: 'labels', range: 'B1:X4', labels: { top: true } },
    { id: 'left-twice', book: 'labels', range: 'A2:C4', labels: { left: true } },
    { id: 'top-right', book: 'labels', range: 'B1:E4', labels: { top: true, right: true } },
  ]
}
