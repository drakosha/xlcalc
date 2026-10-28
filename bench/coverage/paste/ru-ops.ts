/**
 * Paste Special in Excel ru-RU (`paste ru-ops ru`), predicted by the engine before Excel ran them
 * (`ru-ops.predicted.json`): numbers and formulas with an operation, the formula each gives read in Russian
 * (`FormulaLocal`: a decimal comma, `;` between arguments, Russian function names), numbers with formats that hold
 * separators pasted with values and number formats, formats, and an operation, and dates. The targets' formulas go to
 * Excel as a typed formula is saved (`plain`).
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import type { PasteCase, PasteCases, PasteOperation, PasteWhat } from './cases.js'

/** S: a source A1:A6 (numbers with formats of their own, a decimal, a formula, a date, text); targets C1:C6 of numbers and formulas, E1:E6 of numbers with formats. */
function ruBook(): BookInput {
  const cells: Record<string, CellInput> = {
    A1: { value: 1234.5, numberFormat: '#,##0.00' },
    A2: { value: 0.5 },
    A3: { formula: 'IF(B1>1,B1*1.5,0)' },
    A4: { value: 0.125, numberFormat: '0.0%' },
    A5: { value: 46037, numberFormat: 'dd/mm/yyyy' },
    A6: { value: 'текст' },
    B1: { value: 2 },
    C1: { value: 10 },
    C2: { formula: 'SUM(B1:B2)' },
    C3: { value: 2.5 },
    C4: { formula: 'B1/4' },
    C5: { value: 1 },
    C6: { formula: 'ROUND(B1,1)' },
  }
  for (let r = 1; r <= 6; r++) cells[`E${r}`] = { value: 1000 + r, numberFormat: '#,##0.000' }
  return { locale: 'ru-RU', sheets: [{ name: 'S', cells }] }
}

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, to: string, what: PasteWhat, o: { operation?: PasteOperation; local?: string[] } = {}) => cases.push({ id, book: 'ru', from: { sheet: 'S', range: 'A1:A6' }, to: { sheet: 'S', range: to }, what, ...o })
  const formulas = ['C2', 'C3', 'C4', 'C6']
  for (const operation of ['add', 'subtract', 'multiply', 'divide'] as const) add(`formulas-${operation}`, 'C1', 'formulas', { operation, local: formulas })
  // a value pasted onto a value with an operation gives a value: C3 is read with the formulas only
  add('values-add', 'C1', 'values', { operation: 'add', local: ['C2', 'C4', 'C6'] })
  add('all-add', 'C1', 'all', { operation: 'add', local: formulas })
  add('values-number-formats', 'E1', 'valuesAndNumberFormats')
  add('formats', 'E1', 'formats')
  add('values-add-formats', 'E1', 'valuesAndNumberFormats', { operation: 'add' })
  add('formulas-number-formats-multiply', 'E1', 'formulasAndNumberFormats', { operation: 'multiply', local: ['E3'] })
  return { books: { ru: ruBook() }, cases, plain: { ru: ['A3', 'C2', 'C4', 'C6'] } }
}
