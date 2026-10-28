/**
 * The Format Painter of Excel for Mac against Paste Special's Formats: each case twice, painted and pasted with Formats.
 * One cell over a range, a range onto one cell, a merged area, whole rows (their height and look) and whole columns
 * (their width and look), onto a merged area, from another sheet; the source holds a conditional format, data
 * validation and an unlocked cell. The engine predicts Formats for both (`painter.predicted.json`).
 */
import type { BookInput, CellStyle } from '../../../src/contract.js'
import { RED } from '../cond/cases.js'
import type { PasteCase, PasteCases, PasteWhat } from './cases.js'

const THIN = { style: 'thin' } as const
const STYLES: CellStyle[] = [
  {},
  { font: { bold: true }, fill: { color: '#FFFF00' }, border: { bottom: THIN } },
  { font: { italic: true }, protection: { locked: false } },
  { fill: { color: '#9BC2E6' } },
  { font: { size: 16 } },
]

/** S: the source A1:B2 (a rule and a validation over it, B2 unlocked), A4:B4 merged, row 6 and column D of their own; targets with values. */
function painterBook(): BookInput {
  return {
    styles: STYLES,
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { value: 1, style: 1, numberFormat: '0.00' },
          B1: { value: 2, style: 2 },
          A2: { value: 3, numberFormat: '0%' },
          B2: { value: 4, style: 2 },
          A4: { value: 'm', style: 3 },
          A6: { value: 'row' },
          D8: { value: 'col' },
          F3: { value: 'x' },
          G4: { value: 'y', style: 3 },
          F10: { value: 'z' },
          F12: { value: 'w' },
          H14: { value: 'v' },
          J2: { value: 'j' },
          A8: { value: 'r8' },
        },
        merges: ['A4:B4', 'H14:I14'],
        rows: { 6: { height: 30, style: 4 } },
        columns: { D: { width: 25, style: 3 } },
        conditionalFormats: [{ type: 'cellIs', ranges: ['A1:B2'], operator: 'greaterThan', formulas: ['1'], style: RED }],
        validations: [{ range: 'A1:B2', type: 'whole', operator: 'between', formula1: '0', formula2: '10' }],
      },
      { name: 'T', cells: { B2: { value: 't' } } },
    ],
  }
}

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, from: string, to: string, toSheet = 'S') => {
    for (const what of ['painter', 'formats'] as PasteWhat[]) cases.push({ id: `${id}-${what}`, book: 'painter', from: { sheet: 'S', range: from }, to: { sheet: toSheet, range: to }, what })
  }
  add('cell-to-range', 'A1', 'F3:H5')
  add('range-to-cell', 'A1:B2', 'F10')
  add('merged', 'A4:B4', 'F12')
  // the merged area whole: the painter's selection with one more column shrinks onto it, where from H14 it grew to G14
  add('onto-merged', 'A1', 'H14:I14')
  add('row', '6:6', '8:8')
  add('column', 'D:D', 'J:J')
  add('other-sheet', 'A1:B2', 'B2', 'T')
  return { books: { painter: painterBook() }, cases }
}
