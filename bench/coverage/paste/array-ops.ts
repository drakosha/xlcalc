/**
 * Paste Special's operations onto a legacy array formula (Ctrl+Shift+Enter) in both Excels: onto its whole range with
 * values, formulas and All, from a copy laid at its corner, onto a one-cell array, a two-by-two one, the whole range and
 * one cell more, part of the range; values without an operation and with Skip Blanks, and Formats with an operation, onto
 * the whole range. Predicted by the engine before Excel ran them (`array-ops.predicted.json`).
 */
import type { BookInput } from '../../../src/contract.js'
import type { PasteCase, PasteCases } from './cases.js'

/** S: numbers in A1:B3, the legacy arrays C1:C2, E1 and G1:H2. */
function arraysBook(): BookInput {
  return {
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { value: 2 },
          A2: { value: 3 },
          A3: { value: 4 },
          B1: { value: 5 },
          B2: { value: 6 },
          C1: { formula: 'A1:A2*10', array: 'C1:C2' },
          E1: { formula: 'SUM(A1:A2)', array: 'E1' },
          G1: { formula: 'A1:B2+1', array: 'G1:H2' },
        },
      },
    ],
  }
}

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, from: string, to: string, what: PasteCase['what'], o: Partial<Pick<PasteCase, 'operation' | 'skipBlanks'>> = {}) =>
    cases.push({ id, book: 'arrays', from: { sheet: 'S', range: from }, to: { sheet: 'S', range: to }, what, ...o })
  add('whole-add-values', 'A1:A2', 'C1:C2', 'values', { operation: 'add' })
  add('whole-add-formulas', 'A1:A2', 'C1:C2', 'formulas', { operation: 'add' })
  add('whole-add-all', 'A1:A2', 'C1:C2', 'all', { operation: 'add' })
  add('corner-multiply', 'A1:A2', 'C1', 'values', { operation: 'multiply' })
  add('one-cell-add', 'A1', 'E1', 'values', { operation: 'add' })
  add('one-cell-add-all', 'A1', 'E1', 'all', { operation: 'add' })
  add('square-add', 'A1:B2', 'G1', 'values', { operation: 'add' })
  add('beyond-add', 'A1:A3', 'C1', 'values', { operation: 'add' })
  add('part-add', 'A1', 'C1', 'values', { operation: 'add' })
  add('whole-values', 'A1:A2', 'C1:C2', 'values')
  add('whole-skip', 'A1:A2', 'C1:C2', 'values', { skipBlanks: true })
  add('formats-add', 'A1:A2', 'C1:C2', 'formats', { operation: 'add' })
  return { books: { arrays: arraysBook() }, cases }
}
