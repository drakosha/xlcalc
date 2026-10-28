/**
 * The first look at Paste Special's operations and Skip Blanks, before the engine has them: a column of a number, text,
 * a blank, an error, a formula, a date, a zero and a boolean added to, subtracted from, multiplied into and divided into
 * numbers; a number, a formula and text added onto such a column; an operation with every kind, transposed and over
 * merged areas; Skip Blanks with the kinds, with an operation, transposed and over a blank that has a look. The engine's
 * prediction (`ops.predicted.json`) has none of them.
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import type { PasteCase, PasteCases, PasteOperation, PasteWhat } from './cases.js'

/** The source A1:A9 (A9 a blank with a look), numbers C1:C9 to paste onto, a mixed column D1:D8, sources E1:E3, merged areas, a row L1:S1 of numbers. */
function opsBook(): BookInput {
  const cells: Record<string, CellInput> = {
    A1: { value: 2 },
    A2: { value: 'txt' },
    A4: { formula: '1/0' },
    A5: { formula: 'B1*2', style: 1 },
    A6: { value: 46037, numberFormat: 'm/d/yyyy' },
    A7: { value: 0 },
    A8: { value: true },
    A9: { style: 1 },
    B1: { value: 3 },
    B2: { value: 4 },
    D1: { value: 'text' },
    D3: { formula: 'B1+1' },
    D4: { formula: '1/0' },
    D5: { value: 46037, numberFormat: 'm/d/yyyy' },
    D6: { value: true },
    D7: { value: 10, note: { text: 'n' } },
    D8: { formula: 'SUM(B1:B2)' },
    E1: { value: 5, numberFormat: '0%' },
    E2: { formula: 'B1*2' },
    E3: { value: 'txt' },
    G1: { value: 10 },
    I1: { value: 2 },
  }
  for (let r = 1; r <= 9; r++) cells[`C${r}`] = { value: 10, ...(r === 1 ? { numberFormat: '0.0' } : {}) }
  for (const c of 'LMNOPQRS') cells[`${c}1`] = { value: 10 }
  return { styles: [{}, { font: { bold: true } }], sheets: [{ name: 'S', cells, merges: ['G1:H1', 'I1:J1'] }] }
}

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, from: string, to: string, what: PasteWhat, o: { operation?: PasteOperation; skipBlanks?: boolean; transpose?: boolean }) =>
    cases.push({ id, book: 'ops', from: { sheet: 'S', range: from }, to: { sheet: 'S', range: to }, what, ...o })
  for (const operation of ['add', 'subtract', 'multiply', 'divide'] as const) add(`${operation}-values`, 'A1:A8', 'C1', 'values', { operation })
  add('add-all', 'A1:A8', 'C1', 'all', { operation: 'add' })
  add('add-formulas', 'A1:A8', 'C1', 'formulas', { operation: 'add' })
  add('divide-all', 'A1:A8', 'C1', 'all', { operation: 'divide' })
  add('number-add-all', 'E1', 'D1:D8', 'all', { operation: 'add' })
  add('number-add-values', 'E1', 'D1:D8', 'values', { operation: 'add' })
  add('formula-add-formulas', 'E2', 'D1:D8', 'formulas', { operation: 'add' })
  add('formula-add-all', 'E2', 'D1:D8', 'all', { operation: 'add' })
  add('formula-subtract-values', 'E2', 'D1:D8', 'values', { operation: 'subtract' })
  add('text-add-values', 'E3', 'D1:D8', 'values', { operation: 'add' })
  for (const what of ['formats', 'validation', 'notes', 'columnWidths', 'valuesAndNumberFormats', 'formulasAndNumberFormats', 'allExceptBorders'] as const) add(`add-${what}`, 'A1:A8', 'C1', what, { operation: 'add' })
  add('add-transposed', 'A1:A8', 'L1', 'values', { operation: 'add', transpose: true })
  add('add-merged', 'I1:J1', 'G1', 'all', { operation: 'add' })
  add('add-merged-values', 'I1:J1', 'G1', 'values', { operation: 'add' })
  for (const what of ['all', 'values', 'formulas', 'formats', 'notes', 'allExceptBorders'] as const) add(`skip-${what}`, 'A1:A9', 'C1', what, { skipBlanks: true })
  add('skip-add', 'A1:A9', 'C1', 'values', { skipBlanks: true, operation: 'add' })
  add('skip-transposed', 'A1:A9', 'L1', 'all', { skipBlanks: true, transpose: true })
  add('skip-onto-blank', 'A1:A4', 'D1', 'all', { skipBlanks: true })
  return { books: { ops: opsBook() }, cases }
}
