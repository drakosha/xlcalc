/**
 * Formats of rows and columns, en-US, held out: made after the set `files` changed the writer and the reader, and
 * predicted before Excel ran it (`held.predicted.json`). Spans of rows and columns, a row with a style alone crossing
 * columns with a format and a style, a row with a format alone crossing columns with a style, a row with both, dates,
 * a spill across formatted columns, cells with their own General, plain style or italic; the file as written and
 * without the cells where lines cross, values typed into both; Excel formatting a row and then columns crossing it.
 */
import type { BookInput, CellStyle } from '../../../src/contract.js'
import type { LineStep, LinesCase, LinesCases } from './cases.js'

const STYLES: CellStyle[] = [
  { font: { italic: true } },
  { fill: { color: '#C6EFCE' } },
  { font: { bold: true, color: '#9C0006' } },
  { fill: { color: '#DDEBF7' }, font: { bold: true } },
  {},
]

/**
 * Columns: B:C a date, D a fill, E:F a number format and a bold red font, H a fill and bold. Rows: 3:4 italic alone,
 * 6 a percent alone, 9 a fill and a format, 11 a date alone.
 */
function grid(): BookInput {
  return {
    styles: STYLES,
    sheets: [
      {
        name: 'H',
        cells: {
          A1: { value: 'held' },
          B2: { value: { date: '2025-01-31' } },
          C2: { formula: 'B2+30' },
          B3: { value: 46096 },
          D3: { value: 7 },
          H4: { value: 'x' },
          B6: { value: 0.25 },
          E6: { value: 1.5 },
          B9: { value: 3 },
          E9: { value: 2.25 },
          E11: { value: 46096 },
          B13: { formula: 'SEQUENCE(1,6,1)' },
          A15: { value: 5, format: 'General' },
          D15: { value: 1, style: 4 },
          E15: { style: 0 },
          F15: { value: 46096, format: 'General', style: 1 },
        },
        columns: { 'B:C': { format: 'yyyy-mm-dd', width: 14 }, D: { style: 1, width: 12 }, 'E:F': { format: '#,##0.0', style: 2, width: 12 }, H: { style: 3, width: 12 } },
        rows: { '3:4': { style: 0 }, 6: { format: '0.0%' }, 9: { format: '0', style: 1 }, 11: { format: 'yyyy-mm-dd' } },
      },
    ],
  }
}

const blank = (): BookInput => ({ sheets: [{ name: 'H', cells: { A1: { value: 'x' } }, columns: { 'B:F': { width: 16 } } }] })

const READ: [string, string][] = [['H', 'A1:H15']]

/** Values typed where lines cross. */
const TYPED: LineStep[] = [
  { sheet: 'H', range: 'C3', value: 46096 },
  { sheet: 'H', range: 'E3', value: 12.5 },
  { sheet: 'H', range: 'H3', value: 'y' },
  { sheet: 'H', range: 'D6', value: 0.5 },
  { sheet: 'H', range: 'F6', value: 0.5 },
  { sheet: 'H', range: 'D11', value: 46096 },
  { sheet: 'H', range: 'G6', value: 3 },
]

export function buildLinesCases(): LinesCases {
  const cases: LinesCase[] = [
    { id: 'written', book: 'grid', read: READ },
    { id: 'bare', book: 'grid', bare: true, read: READ },
    { id: 'typed', book: 'grid', steps: TYPED, read: READ },
    { id: 'bare-typed', book: 'grid', bare: true, steps: TYPED, read: READ },
    // a row's fill, then a column's number format crossing it; a row's italic, then a column's fill crossing it
    {
      id: 'excel-fill-format',
      book: 'blank',
      steps: [
        { sheet: 'H', range: '5:5', fill: '#DDEBF7' },
        { sheet: 'H', range: 'C:D', format: '0.000' },
        { sheet: 'H', range: 'C5', value: 2.5 },
        { sheet: 'H', range: 'E5', value: 2.5 },
        { sheet: 'H', range: 'C6', value: 2.5 },
      ],
      read: [['H', 'B4:E6']],
    },
    {
      id: 'excel-italic-fill',
      book: 'blank',
      steps: [
        { sheet: 'H', range: '7:7', italic: true },
        { sheet: 'H', range: 'E:E', fill: '#C6EFCE' },
        { sheet: 'H', range: 'E7', value: 'z' },
      ],
      read: [['H', 'D6:F8']],
    },
  ]
  return { books: { grid: grid(), blank: blank() }, cases }
}
