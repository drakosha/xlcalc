/**
 * Formats of rows, columns and rectangles as a session sets them (`patchCells`), en-US, held out: the engine's answers
 * are those of the steps taken by a session (`sessionCase`), predicted before Excel ran the set (`patch.predicted.json`).
 * A rectangle over a filled column and across a bold row, its empty cells too; a column over cells with a style or a
 * format of their own; a row's date format across a column of percents; a column filled, then bold, then a row italic
 * across it; a row italic over a filled cell; a column filled across a row with a number format alone; a style set over
 * a book's own column and row styles.
 */
import type { BookInput } from '../../../src/contract.js'
import type { LinesCase, LinesCases } from './cases.js'

const blank = (): BookInput => ({ sheets: [{ name: 'P', cells: { A1: { value: 'x' } }, columns: { 'B:H': { width: 14 } } }] })

/** A column with a fill and a percent format, a row bold with a date, written in the book. */
const lined = (): BookInput => ({
  styles: [{ fill: { color: '#C6EFCE' } }, { font: { bold: true } }],
  sheets: [
    {
      name: 'P',
      cells: { A1: { value: 'x' }, C2: { value: 0.25 }, C5: { value: 46096 }, E5: { value: 7 } },
      columns: { C: { style: 0, format: '0%', width: 14 }, 'D:F': { width: 14 } },
      rows: { 5: { style: 1, format: 'yyyy-mm-dd' } },
    },
  ],
})

export function buildLinesCases(): LinesCases {
  const cases: LinesCase[] = [
    {
      id: 'rect-over-column',
      book: 'blank',
      steps: [
        { sheet: 'P', range: 'B:B', fill: '#FFFF00' },
        { sheet: 'P', range: '3:3', italic: true },
        { sheet: 'P', range: 'B2:C4', bold: true },
        { sheet: 'P', range: 'B2', value: 5 },
        { sheet: 'P', range: 'C3', value: 6 },
      ],
      read: [['P', 'A1:D5']],
    },
    {
      id: 'column-over-cells',
      book: 'blank',
      steps: [
        { sheet: 'P', range: 'C3', italic: true },
        { sheet: 'P', range: 'C5', format: '0%' },
        { sheet: 'P', range: 'C3', value: 1.5 },
        { sheet: 'P', range: 'C5', value: 0.5 },
        { sheet: 'P', range: 'C6', value: 2 },
        { sheet: 'P', range: 'C:C', bold: true, format: '0.00' },
        { sheet: 'P', range: 'C7', value: 3 },
      ],
      read: [['P', 'B2:D8']],
    },
    {
      id: 'row-date-across-percent',
      book: 'blank',
      steps: [
        { sheet: 'P', range: 'D:D', format: '0%' },
        { sheet: 'P', range: '4:4', format: 'yyyy-mm-dd' },
        { sheet: 'P', range: 'D4', value: 46096 },
        { sheet: 'P', range: 'D5', value: 0.5 },
        { sheet: 'P', range: 'E4', value: 46096 },
      ],
      read: [['P', 'C3:E5']],
    },
    {
      id: 'column-twice-row-across',
      book: 'blank',
      steps: [
        { sheet: 'P', range: 'E:E', fill: '#DDEBF7' },
        { sheet: 'P', range: 'E:E', bold: true },
        { sheet: 'P', range: '6:6', italic: true },
        { sheet: 'P', range: 'E2', value: 1 },
        { sheet: 'P', range: 'E6', value: 2 },
        { sheet: 'P', range: 'F6', value: 3 },
      ],
      read: [['P', 'D1:F7']],
    },
    {
      id: 'row-over-filled-cell',
      book: 'blank',
      steps: [
        { sheet: 'P', range: 'C7', fill: '#FFC7CE' },
        { sheet: 'P', range: '7:7', italic: true },
        { sheet: 'P', range: 'C7', value: 'a' },
        { sheet: 'P', range: 'D7', value: 'b' },
      ],
      read: [['P', 'B6:D8']],
    },
    {
      id: 'column-fill-across-row-format',
      book: 'blank',
      steps: [
        { sheet: 'P', range: '8:8', format: '0.000' },
        { sheet: 'P', range: 'F:F', fill: '#FFFF00' },
        { sheet: 'P', range: 'F8', value: 1.5 },
        { sheet: 'P', range: 'G8', value: 1.5 },
        { sheet: 'P', range: 'F9', value: 1.5 },
      ],
      read: [['P', 'E7:G9']],
    },
    {
      id: 'over-lines-of-the-book',
      book: 'lined',
      steps: [
        { sheet: 'P', range: 'C:C', italic: true },
        { sheet: 'P', range: 'D4:E6', fill: '#FFC7CE' },
        { sheet: 'P', range: '5:5', format: '0.0' },
        { sheet: 'P', range: 'D5', value: 2.5 },
      ],
      read: [['P', 'B1:F7']],
    },
  ]
  return { books: { blank: blank(), lined: lined() }, cases, session: true }
}
