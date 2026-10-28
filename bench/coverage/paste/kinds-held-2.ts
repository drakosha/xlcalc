/**
 * Held out after the set `kinds-held` fitted the engine's column widths of a transposed paste, the notes a paste of
 * notes lays over merged areas and what Skip Blanks takes for blank: predicted by the engine before Excel ran them
 * (`kinds-held-2.predicted.json`). Widths of a source of two rows and of one row turned and tiled; notes of two rows over
 * a merged area of two by two; Skip Blanks over a cell holding a note alone with All, All except Borders and Formats.
 */
import type { BookInput } from '../../../src/contract.js'
import type { PasteCase, PasteCases, PasteWhat } from './cases.js'

/** A source A1:C2 (columns A 20, B 5, C 9; notes in B1, A2 and C1 alone); targets of their own, F10:G11 merged. */
function heldBook(): BookInput {
  return {
    styles: [{}, { font: { bold: true } }],
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { value: 1, style: 1 },
          B1: { value: 2, note: { text: 'nb' } },
          C1: { note: { text: 'alone' } },
          A2: { value: 3, note: { text: 'na2' } },
          F2: { value: 'f' },
          H2: { value: 'h', style: 1, note: { text: 'own-h' } },
          F10: { value: 'x', note: { text: 'own' } },
        },
        merges: ['F10:G11'],
        columns: { A: { width: 20 }, B: { width: 5 }, C: { width: 9 }, L: { width: 30 } },
      },
    ],
  }
}

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, from: string, to: string, what: PasteWhat, o: { transpose?: boolean; skipBlanks?: boolean } = {}) =>
    cases.push({ id, book: 'held', from: { sheet: 'S', range: from }, to: { sheet: 'S', range: to }, what, ...o })
  add('columnWidths-two-rows-transposed', 'A1:C2', 'K1', 'columnWidths', { transpose: true })
  add('columnWidths-row-transposed-tiled', 'A1:B1', 'K1:N2', 'columnWidths', { transpose: true })
  add('notes-two-rows-merged', 'A1:B2', 'F10', 'notes')
  add('notes-merged-none', 'A1', 'F10', 'notes')
  add('all-skip-note-alone', 'A1:C1', 'F2', 'all', { skipBlanks: true })
  add('allExceptBorders-skip-note-alone', 'A1:C1', 'F2', 'allExceptBorders', { skipBlanks: true })
  add('formats-skip-note-alone', 'A1:C1', 'F2', 'formats', { skipBlanks: true })
  return { books: { held: heldBook() }, cases }
}
