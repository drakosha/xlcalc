/**
 * Held out after the set `ops` gave the engine Paste Special's operations and Skip Blanks: cases it did not have,
 * predicted by the engine before Excel ran them (`ops-held.predicted.json`). The book's formulas but the spill go to
 * Excel as it saves a typed formula (`plain`), not as dynamic array formulas, which it refuses to combine (the set
 * `ops`): a number, a formula and a decimal combined with text, a blank, formulas, a date, a boolean, decimals and
 * negative numbers; two blanks; an operation onto a spill, onto another sheet and onto a merged area the copy does not
 * bring; the look of a text pasted with an operation; Skip Blanks over a formula giving empty text.
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import type { PasteCase, PasteCases, PasteOperation, PasteWhat } from './cases.js'

/** Targets D1:D11 of every kind, D12 a spill, sources E1:E4, a blank F1, merged F20:G20; T a target on another sheet. */
function heldBook(): BookInput {
  const cells: Record<string, CellInput> = {
    B1: { value: 3 },
    B2: { value: 4 },
    D1: { value: 'text' },
    D3: { formula: 'B1+1' },
    D4: { formula: 'NA()' },
    D5: { value: 46037, numberFormat: 'm/d/yyyy' },
    D6: { value: true },
    D7: { value: 10, note: { text: 'n' } },
    D8: { formula: 'SUM(B1:B2)' },
    D9: { value: 1.5 },
    D10: { value: -3 },
    D11: { value: 0.1 },
    D12: { formula: 'SEQUENCE(2)' },
    E1: { value: 5, numberFormat: '0%' },
    E2: { formula: 'B1*2' },
    E3: { value: 0.2 },
    E4: { value: 'txt', style: 1 },
    E5: { formula: '""' },
    F20: { value: 1 },
  }
  return {
    styles: [{}, { font: { bold: true } }],
    sheets: [
      { name: 'S', cells, merges: ['F20:G20'] },
      { name: 'T', cells: { A1: { value: 7 }, A2: { formula: 'A1*3' } } },
    ],
  }
}

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, from: string, to: string, what: PasteWhat, o: { operation?: PasteOperation; skipBlanks?: boolean }, toSheet = 'S') =>
    cases.push({ id, book: 'held', from: { sheet: 'S', range: from }, to: { sheet: toSheet, range: to }, what, ...o })
  add('number-add-all', 'E1', 'D1:D11', 'all', { operation: 'add' })
  add('number-add-values', 'E1', 'D1:D11', 'values', { operation: 'add' })
  add('number-divide-values', 'E1', 'D1:D11', 'values', { operation: 'divide' })
  add('formula-add-formulas', 'E2', 'D1:D11', 'formulas', { operation: 'add' })
  add('formula-multiply-all', 'E2', 'D1:D11', 'all', { operation: 'multiply' })
  add('formula-subtract-values', 'E2', 'D1:D11', 'values', { operation: 'subtract' })
  add('decimal-add-values', 'E3', 'D1:D11', 'values', { operation: 'add' })
  add('blank-multiply-blank', 'F1', 'D2', 'values', { operation: 'multiply' })
  add('blank-divide-blank', 'F1', 'D2', 'all', { operation: 'divide' })
  add('number-subtract-blank', 'E1', 'D2', 'values', { operation: 'subtract' })
  add('number-add-spill', 'E1', 'D12', 'values', { operation: 'add' })
  add('number-add-other-sheet', 'E1:E2', 'A1', 'formulas', { operation: 'add' }, 'T')
  add('number-add-merged', 'E1:E2', 'F20', 'values', { operation: 'add' })
  add('text-add-all', 'E4', 'D9', 'all', { operation: 'add' })
  add('empty-text-skip', 'E5', 'D9', 'all', { skipBlanks: true })
  add('empty-text-values-skip', 'E5', 'D9', 'values', { skipBlanks: true })
  // the spill D12 stays a dynamic array formula
  return { books: { held: heldBook() }, cases, plain: { held: ['A2', 'D3', 'D4', 'D8', 'E2'] } }
}
