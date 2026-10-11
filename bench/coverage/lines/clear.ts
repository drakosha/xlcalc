/**
 * Cells cleared as a session clears them (`clearCells`), en-US, held out: Excel's ClearContents, ClearFormats and Clear
 * over rectangles, whole columns and whole rows of a book with styles, number formats, notes, hyperlinks, a formatted
 * column, a bold row, conditional formats and rules of data validation. The engine's answers are those of the steps
 * taken by a session (`sessionCase`, `partsOf`), predicted before Excel ran the set (`clear.predicted.json`): what the
 * read cells show, and the notes, hyperlinks, conditional formats and rules of the book Excel saved.
 */
import type { LinesCase, LinesCases, LineStep } from './cases.js'
import type { StoredBook } from '../stored.js'

/**
 * Cells with a style, a format, a note and a hyperlink in B2:D4; column F filled with a number format, with a cell of its
 * own italic; row 8 bold. A rule over B2:D4 written for B2, one over F:F; rules of data validation over B2:C5 and F2:F10.
 */
function book(): StoredBook {
  return {
    styles: [{ font: { bold: true } }, { fill: { color: '#FFFF00' } }, { font: { italic: true } }, { fill: { color: '#DDEBF7' } }],
    sheets: [
      {
        name: 'C',
        cells: {
          A1: { value: 'x' },
          B2: { value: 5, style: 0, format: '0.00', note: { author: 'A', text: 'b2' }, link: { url: 'https://example.com/b2' } },
          B3: { formula: 'B2*2', style: 1 },
          C2: { value: 'a', note: { author: 'A', text: 'c2' } },
          C3: { value: 7, link: { location: "'C'!A1" } },
          D2: { value: 1, format: '0%' },
          D4: { value: 4, style: 3 },
          F2: { value: 2 },
          F3: { value: 4, style: 2 },
          F5: { value: 6, note: { author: 'A', text: 'f5' } },
          G8: { value: 1 },
          E8: { value: 2, style: 3 },
        },
        columns: { 'B:E': { width: 12 }, F: { style: 1, format: '0.0', width: 12 }, G: { width: 12 } },
        rows: { 8: { style: 0 } },
        conditionalFormats: [
          { type: 'expression', ranges: ['B2:D4'], formula: 'B2>3', style: { font: { color: '#C00000' } } },
          { type: 'cellIs', ranges: ['F:F'], operator: 'greaterThan', formulas: ['3'], style: { font: { color: '#0070C0' } } },
        ],
        validations: [
          { range: 'B2:C5', type: 'whole', operator: 'greaterThan', formula1: '0' },
          { range: 'F2:F10', type: 'decimal', operator: 'lessThan', formula1: 'F1+100' },
        ],
      },
    ],
  }
}

const READ: [string, string][] = [['C', 'A1:G9']]

const clear = (range: string, what: LineStep['clear']): LinesCase => ({ id: `${what}-${range.replace(':', '-')}`, book: 'book', steps: [{ sheet: 'C', range, clear: what }], read: READ, parts: true })

export function buildLinesCases(): LinesCases {
  const cases: LinesCase[] = [
    // a rectangle of cells with values, formulas, styles, formats, notes and hyperlinks, the top of a rule's cells
    clear('B2:C3', 'contents'),
    clear('B2:C3', 'formats'),
    clear('B2:C3', 'all'),
    clear('B2:D2', 'formats'),
    // cells of a formatted column, one with a style of its own, one empty
    clear('F3:F4', 'formats'),
    clear('F3:F4', 'all'),
    // a whole column with a look, crossing a bold row; a whole row with a look
    clear('F:F', 'contents'),
    clear('F:F', 'formats'),
    clear('F:F', 'all'),
    clear('8:8', 'formats'),
    clear('8:8', 'all'),
    // a value typed into a cell after its formats were cleared, in the formatted column and the bold row
    {
      id: 'formats-then-typed',
      book: 'book',
      steps: [
        { sheet: 'C', range: 'F2:G8', clear: 'formats' },
        { sheet: 'C', range: 'F4', value: 3.25 },
        { sheet: 'C', range: 'G8', value: 3.25 },
      ],
      read: READ,
      parts: true,
    },
  ]
  return { books: { book: book() }, cases, session: true }
}
