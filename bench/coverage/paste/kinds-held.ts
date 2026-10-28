/**
 * Held out after the set `kinds` gave the engine Paste Special's other kinds: cases it did not have, predicted by the
 * engine before Excel ran them (`kinds-held.predicted.json`). Column widths of a source that is not square, transposed;
 * values and number formats of whole columns with a format of their own and of a spill; formulas and number formats
 * onto merged areas; all except borders where both sides have borders, and of whole columns; notes with Skip Blanks over
 * a cell holding a note alone.
 */
import type { BookInput, CellStyle } from '../../../src/contract.js'
import type { PasteCase, PasteCases, PasteWhat } from './cases.js'

const THIN = { style: 'thin' } as const
const STYLES: CellStyle[] = [{}, { font: { italic: true }, border: { top: THIN, bottom: THIN, left: THIN, right: THIN } }, { fill: { color: '#9BC2E6' }, border: { top: { style: 'thick' } } }, { font: { bold: true } }]

/** A source A1:C1 with widths of its own (A 20, B 5), column E of its own format and look, a spill, notes; targets of their own. */
function heldBook(): BookInput {
  return {
    styles: STYLES,
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { value: 1, style: 1, numberFormat: '0.00' },
          B1: { value: 2, note: { text: 'nb' } },
          C1: { note: { text: 'alone' } },
          E1: { value: 4 },
          E2: { value: 5, numberFormat: '0%' },
          A3: { formula: 'SEQUENCE(2)', numberFormat: '0.0' },
          G2: { value: 'g', style: 2, note: { text: 'own' } },
          H2: { value: 'h', note: { text: 'own-h' } },
          F5: { value: 'x' },
          F7: { value: 'y', numberFormat: '0.000' },
        },
        merges: ['F5:G5', 'F7:H7'],
        columns: { A: { width: 20 }, B: { width: 5 }, E: { width: 12, style: 3, numberFormat: '0.0' } },
      },
    ],
  }
}

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, from: string, to: string, what: PasteWhat, o: { transpose?: boolean; skipBlanks?: boolean } = {}) =>
    cases.push({ id, book: 'held', from: { sheet: 'S', range: from }, to: { sheet: 'S', range: to }, what, ...o })
  add('columnWidths-row-transposed', 'A1:C1', 'K1', 'columnWidths', { transpose: true })
  add('columnWidths-column-transposed', 'A1:A3', 'K1', 'columnWidths', { transpose: true })
  add('columnWidths-onto-merged', 'A1:B1', 'G7', 'columnWidths')
  add('valuesAndNumberFormats-column', 'E:E', 'J:J', 'valuesAndNumberFormats')
  add('formulasAndNumberFormats-column', 'E:E', 'J:J', 'formulasAndNumberFormats')
  add('valuesAndNumberFormats-spill', 'A3:A4', 'L3', 'valuesAndNumberFormats')
  add('formulasAndNumberFormats-merged', 'A1:B1', 'F5', 'formulasAndNumberFormats')
  add('formulasAndNumberFormats-merged-part', 'A1', 'G7', 'formulasAndNumberFormats')
  add('valuesAndNumberFormats-merged-wide', 'A1:C1', 'F7', 'valuesAndNumberFormats')
  add('allExceptBorders-borders', 'A1', 'G2', 'allExceptBorders')
  add('allExceptBorders-whole-column', 'E:E', 'J:J', 'allExceptBorders')
  add('allExceptBorders-transposed', 'A1:C1', 'G2', 'allExceptBorders', { transpose: true })
  add('notes-skip', 'A1:C1', 'F2', 'notes', { skipBlanks: true })
  add('notes-alone', 'C1', 'H2', 'notes')
  add('notes-onto-merged-wide', 'A1:C1', 'F7', 'notes')
  add('valuesAndNumberFormats-skip', 'A1:C1', 'F2', 'valuesAndNumberFormats', { skipBlanks: true })
  return { books: { held: heldBook() }, cases }
}
