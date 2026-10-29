/**
 * Held out after `array-ops` gave the engine its rule (an operation onto any legacy array formula is refused, of one cell
 * too): a copy tiled over a one-cell array and other cells, an operation transposed onto a range with one, with Skip
 * Blanks, over a destination that meets a range's last cell alone, onto another sheet's one-cell array, a paste of
 * formulas with an operation beside an array, which touches none, and a whole array copied and combined onto plain
 * cells. Predicted by the engine before Excel ran them (`array-ops-held.predicted.json`).
 */
import type { BookInput } from '../../../src/contract.js'
import type { PasteCase, PasteCases } from './cases.js'

/** S: numbers in A1:B3, the legacy arrays C1:C2 and E1, plain numbers in J1:J3; T: a one-cell array at B2. */
function heldBook(): BookInput {
  return {
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { value: 2 },
          A2: { value: 3 },
          A3: { value: 4 },
          B1: { value: 5 },
          C1: { formula: 'A1:A2*10', array: 'C1:C2' },
          E1: { formula: 'SUM(A1:A2)', array: 'E1' },
          E2: { value: 7 },
          J1: { value: 1 },
          J2: { value: 1 },
          J3: { value: 1 },
        },
      },
      { name: 'T', cells: { B2: { formula: 'MAX(S!A1:A3)', array: 'B2' } } },
    ],
  }
}

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, from: string, to: string, what: PasteCase['what'], o: Partial<Pick<PasteCase, 'operation' | 'skipBlanks' | 'transpose'>> = {}, toSheet = 'S') =>
    cases.push({ id, book: 'held', from: { sheet: 'S', range: from }, to: { sheet: toSheet, range: to }, what, ...o })
  add('tiled-over-one-cell', 'A1', 'E1:E2', 'values', { operation: 'add' })
  add('transposed-onto-one-cell', 'A1:B1', 'E1', 'values', { operation: 'subtract', transpose: true })
  add('skip-onto-one-cell', 'A1', 'E1', 'all', { operation: 'add', skipBlanks: true })
  add('last-cell-alone', 'A1', 'C2', 'values', { operation: 'multiply' })
  add('other-sheet-one-cell', 'A1', 'B2', 'values', { operation: 'divide' }, 'T')
  add('beside-array', 'A1:A2', 'D1', 'formulas', { operation: 'add' })
  add('array-onto-plain', 'C1:C2', 'J1', 'values', { operation: 'add' })
  add('array-formulas-onto-plain', 'C1:C2', 'J2', 'formulas', { operation: 'add' })
  return { books: { held: heldBook() }, cases }
}
