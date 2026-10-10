/**
 * Formats of rows and columns, en-US: how Excel shows the cells of a book `writeXlsx` wrote with formatted columns and
 * rows (numbers, dates, text, booleans, errors, formulas and a spill without a format of their own, cells with their own
 * style or General, empty cells with a style), and where a row with only a style or only a format crosses a column with
 * the other; how Excel stores a value typed into a formatted column and row, and into where a row crosses a column,
 * formatted by Excel in either order or written so. Predicted before Excel ran it (`files.predicted.json`).
 */
import type { BookInput, CellStyle } from '../../../src/contract.js'
import type { LinesCase, LinesCases } from './cases.js'

const DATE = 'dddd, mmmm d, yyyy'

const STYLES: CellStyle[] = [
  { font: { bold: true } },
  { fill: { color: '#FFFF00' } },
  { font: { italic: true, color: '#C00000' } },
  {},
  { fill: { color: '#C6EFCE' }, font: { bold: true } },
]

/**
 * Columns: B a date and a fill, C a percent, D an italic red font, E a fill alone, G a percent. Rows: 4 a number format
 * alone, 6 bold alone, 8 both. E4 and G6 are empty where a row with one crosses a column with the other.
 */
function formatted(): BookInput {
  return {
    styles: STYLES,
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { value: 'plain' },
          B2: { value: 46096 },
          B3: { formula: 'B2+1' },
          B4: { value: 46096 },
          B6: { value: 46096 },
          B8: { value: 46096 },
          B10: { value: 'abc' },
          B11: { value: true },
          B12: { formula: 'SEQUENCE(2,1,46096)' },
          B15: { value: 46096, format: 'General' },
          B16: { value: 46096, style: 0 },
          B17: { style: 2 },
          B18: { value: 46096, style: 3 },
          B19: { formula: 'NA()' },
          C2: { value: 0.256 },
          C3: { formula: 'C2*2' },
          C6: { value: 0.5 },
          D2: { value: 5 },
          D3: { value: 'text' },
          D4: { value: 1234.5 },
          D8: { value: 2 },
          F4: { value: 1234.5 },
          F6: { value: 7 },
        },
        columns: { B: { format: DATE, style: 1, width: 30 }, C: { format: '0.00%', width: 12 }, D: { style: 2, width: 12 }, E: { style: 1, width: 12 }, G: { format: '0.00%', width: 14 } },
        rows: { 4: { format: '#,##0.00' }, 6: { style: 0 }, 8: { format: '0.000', style: 4 } },
      },
    ],
  }
}

/** An empty sheet with wide columns, for Excel to format. */
const blank = (): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 'x' } }, columns: { 'B:K': { width: 28 } } }] })

const READ: [string, string][] = [['S', 'A1:G19']]

export function buildLinesCases(): LinesCases {
  const cases: LinesCase[] = [
    { id: 'written', book: 'formatted', read: READ },
    // values typed where a row with only a format crosses a column with only a fill, and the other way round
    {
      id: 'typed-crossing',
      book: 'formatted',
      steps: [
        { sheet: 'S', range: 'E4', value: 46096 },
        { sheet: 'S', range: 'G6', value: 46096 },
        { sheet: 'S', range: 'E6', value: 5 },
        { sheet: 'S', range: 'G8', value: 0.5 },
        { sheet: 'S', range: 'B5', value: 46096 },
      ],
      read: READ,
    },
    // Excel formats a column, then a value is typed into it
    {
      id: 'excel-column',
      book: 'blank',
      steps: [
        { sheet: 'S', range: 'B:B', format: DATE, fill: '#FFFF00' },
        { sheet: 'S', range: 'B2', value: 46096 },
        { sheet: 'S', range: 'B3', value: 'abc' },
      ],
      read: [['S', 'A1:C4']],
    },
    {
      id: 'excel-row',
      book: 'blank',
      steps: [
        { sheet: 'S', range: '4:4', format: '#,##0.00', bold: true },
        { sheet: 'S', range: 'D4', value: 1234.5 },
        { sheet: 'S', range: 'E4', value: 'abc' },
      ],
      read: [['S', 'C3:E5']],
    },
    // a column formatted, then a row crossing it: a value typed where they cross, one beside, one below
    {
      id: 'excel-column-row',
      book: 'blank',
      steps: [
        { sheet: 'S', range: 'F:G', format: DATE },
        { sheet: 'S', range: '6:6', bold: true },
        { sheet: 'S', range: 'F6', value: 46096 },
        { sheet: 'S', range: 'H6', value: 7 },
        { sheet: 'S', range: 'F7', value: 46096 },
      ],
      read: [['S', 'E5:H7']],
    },
    // a row formatted, then a column crossing it
    {
      id: 'excel-row-column',
      book: 'blank',
      steps: [
        { sheet: 'S', range: '8:8', bold: true },
        { sheet: 'S', range: 'H:I', format: DATE, fill: '#FFFF00' },
        { sheet: 'S', range: 'H8', value: 46096 },
        { sheet: 'S', range: 'J8', value: 3 },
        { sheet: 'S', range: 'H9', value: 46096 },
      ],
      read: [['S', 'G7:J9']],
    },
    // a row's format alone, then a column's fill alone crossing it
    {
      id: 'excel-format-fill',
      book: 'blank',
      steps: [
        { sheet: 'S', range: '10:10', format: '0.00' },
        { sheet: 'S', range: 'J:K', fill: '#FFFF00' },
        { sheet: 'S', range: 'J10', value: 5 },
      ],
      read: [['S', 'I9:K11']],
    },
  ]
  return { books: { formatted: formatted(), blank: blank() }, cases }
}
