/**
 * Held out after the set `kinds-held-3` fitted Skip Blanks by layer (a cell of the copy without contents leaves the
 * target's contents and look, one without a note its note): predicted by the engine before Excel ran them
 * (`kinds-held-4.predicted.json`). A value with a note over one with a note, a blank with a look and a note over a cell
 * with a value, a look and a note, transposed, and the same without Skip Blanks.
 */
import type { BookInput } from '../../../src/contract.js'
import type { PasteCase, PasteCases, PasteWhat } from './cases.js'

/** A source A1:B2 (A1 a value with a note, B1 a blank with a look and a note, A2 a value, B2 blank); targets of their own. */
function heldBook(): BookInput {
  const own = (v: string) => ({ value: v, style: 2, note: { text: `own-${v}` } })
  return {
    styles: [{}, { font: { bold: true } }, { fill: { color: '#9BC2E6' } }],
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { value: 1, note: { text: 'na' } },
          B1: { style: 1, note: { text: 'nb' } },
          A2: { value: 2 },
          F1: own('f1'),
          G1: own('g1'),
          F2: own('f2'),
          G2: own('g2'),
        },
      },
    ],
  }
}

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, what: PasteWhat, o: { transpose?: boolean; skipBlanks?: boolean }) => cases.push({ id, book: 'held', from: { sheet: 'S', range: 'A1:B2' }, to: { sheet: 'S', range: 'F1' }, what, ...o })
  for (const what of ['all', 'allExceptBorders', 'notes', 'formats', 'values'] as const) add(`${what}-skip`, what, { skipBlanks: true })
  add('all-skip-transposed', 'all', { skipBlanks: true, transpose: true })
  add('all', 'all', {})
  return { books: { held: heldBook() }, cases }
}
