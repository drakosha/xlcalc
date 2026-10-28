/**
 * Held out after the set `kinds-held-2` fitted Skip Blanks to the notes a paste takes: predicted by the engine before
 * Excel ran them (`kinds-held-3.predicted.json`). A blank without a note over a cell with one, and a note alone over a
 * cell with a value and no note, with All, All except Borders and Comments.
 */
import type { BookInput } from '../../../src/contract.js'
import type { PasteCase, PasteCases, PasteWhat } from './cases.js'

/** A source A1:C1 (B1 blank, C1 a note alone); targets F1:H1 with notes and values of their own. */
function heldBook(): BookInput {
  return {
    styles: [{}, { font: { bold: true } }],
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { value: 1 },
          C1: { note: { text: 'alone' } },
          F1: { value: 'f', note: { text: 'own-f' } },
          G1: { value: 'g', style: 1, note: { text: 'own-g' } },
          H1: { value: 'h', style: 1 },
        },
      },
    ],
  }
}

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, what: PasteWhat) => cases.push({ id, book: 'held', from: { sheet: 'S', range: 'A1:C1' }, to: { sheet: 'S', range: 'F1' }, what, skipBlanks: true })
  add('all-skip', 'all')
  add('allExceptBorders-skip', 'allExceptBorders')
  add('notes-skip', 'notes')
  return { books: { held: heldBook() }, cases }
}
