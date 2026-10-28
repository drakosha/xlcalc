/**
 * Held out after the engine refused an operation onto a spilled array, as Excel refused it in the set `ops-held`:
 * predicted by the engine before Excel ran them (`spill-ops.predicted.json`). An operation onto the spill's formula
 * cell, onto a spilled cell alone, over the whole spill, over part of it and beyond; with values and with formulas;
 * values without an operation onto a spilled cell and Skip Blanks without one over the spill; an operation onto a
 * formula of one cell that is no array, typed (`plain`).
 */
import type { BookInput } from '../../../src/contract.js'
import type { PasteCase, PasteCases, PasteOperation, PasteWhat } from './cases.js'

/** S: sources A1:A3 (A2 blank); D1 `SEQUENCE(3)` spilled over D1:D3; F1 a typed formula. */
function spillBook(): BookInput {
  return { sheets: [{ name: 'S', cells: { A1: { value: 5 }, A3: { value: 7 }, B1: { value: 2 }, D1: { formula: 'SEQUENCE(3)' }, F1: { formula: 'B1*3' } } }] }
}

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, from: string, to: string, what: PasteWhat, o: { operation?: PasteOperation; skipBlanks?: boolean } = {}) => cases.push({ id, book: 'spill', from: { sheet: 'S', range: from }, to: { sheet: 'S', range: to }, what, ...o })
  add('add-formula-cell', 'A1', 'D1', 'values', { operation: 'add' })
  add('add-spilled-cell', 'A1', 'D2', 'values', { operation: 'add' })
  add('multiply-whole', 'A1:A3', 'D1', 'formulas', { operation: 'multiply' })
  add('add-part-beyond', 'A1:A3', 'D2', 'all', { operation: 'add' })
  add('values-spilled-cell', 'A1', 'D2', 'values')
  add('skip-over-spill', 'A1:A3', 'D1', 'all', { skipBlanks: true })
  add('add-typed-formula', 'A1', 'F1', 'values', { operation: 'add' })
  return { books: { spill: spillBook() }, cases, plain: { spill: ['F1'] } }
}
