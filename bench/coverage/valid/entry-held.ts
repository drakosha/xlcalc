/**
 * Held out: whether a value meets its rule, as Excel's `Validation.Value` checks a cell's value, predicted by
 * `Session.validateEntry` before Excel ran it (`entry-held.predicted.json`). One column per rule, row 1 its title, the
 * values below it: text, numbers, booleans, errors and blanks against each type, bounds on blank cells with blanks ignored
 * and not, reversed bounds, numbers equal only to 15 digits, dates with a time, literal lists of numbers, booleans and
 * spaces, lists of cells holding numbers as text.
 */
import type { BookInput, CellInput, DataValidation } from '../../../src/contract.js'
import type { ValidCases } from './cases.js'

type Column = { rule: Omit<DataValidation, 'range'>; values: (CellInput | null)[] }

const v = (value: CellInput['value']): CellInput => ({ value })
const f = (formula: string): CellInput => ({ formula })

/** The rules and their values, by column from A. */
const COLUMNS: Column[] = [
  { rule: { type: 'list', formula1: '"Yes,No,1,2.5,TRUE, spaced,1/2/2025"' }, values: [v('Yes'), v('No'), v(1), f('"1"'), v(2.5), v(true), f('"TRUE"'), v(' spaced'), v('spaced'), v(45659), v('1/2/2025'), null] },
  { rule: { type: 'whole', formula1: '1', formula2: '10' }, values: [v(1), v(10), v(0), v(11), f('5*1'), v(5.5), f('"5"'), v(true), f('NA()'), null, v(1e-20)] },
  { rule: { type: 'whole', formula1: '10', formula2: '1' }, values: [v(5), v(1), v(10), v(0)] },
  { rule: { type: 'decimal', operator: 'notBetween', formula1: '1', formula2: '2' }, values: [v(0.5), v(1), v(1.5), v(2), v(3), v('abc'), v(false)] },
  { rule: { type: 'decimal', operator: 'equal', formula1: '0.3' }, values: [f('0.1+0.2'), v(0.3), v(0.30000000000000004)] },
  { rule: { type: 'date', formula1: 'DATE(2025,1,1)', formula2: 'DATE(2025,1,31)' }, values: [v(45658), v(45688), v(45688.5), v(45689), v(45670.25), f('"2025-01-15"'), v(45657.99)] },
  { rule: { type: 'time', operator: 'lessThan', formula1: '0.5' }, values: [v(0.25), v(0.5), v(1.25), v(45658.25), v(-0.1)] },
  { rule: { type: 'textLength', operator: 'equal', formula1: '3' }, values: [v('abc'), v('ab'), v(123), v(1.5), v(true), v(false), f('0.1+0.2'), v(12.25), f('NA()')] },
  { rule: { type: 'custom', formula1: 'ISNUMBER(I2)' }, values: [v(5), v('x'), f('1/0')] },
  { rule: { type: 'custom', formula1: 'J2*1' }, values: [v(5), v(0), v('x'), v(-1)] },
  { rule: { type: 'custom', formula1: '$Z$1' }, values: [v(5), v('x')] },
  { rule: { type: 'custom', formula1: '$Z$1', allowBlank: false }, values: [v(5), v('x')] },
  { rule: { type: 'custom', formula1: '"TRUE"' }, values: [v(5)] },
  { rule: { type: 'whole', formula1: '$Z$1', formula2: '10' }, values: [v(50), v(5), v(-1)] },
  { rule: { type: 'whole', formula1: '$Z$1', formula2: '10', allowBlank: false }, values: [v(50), v(5), v(-1)] },
  { rule: { type: 'list', formula1: '$Z$1:$Z$3' }, values: [v('anything'), v(1)] },
  { rule: { type: 'list', formula1: '$Z$1:$Z$3', allowBlank: false }, values: [v('anything'), f('""')] },
  { rule: { type: 'list', formula1: '$Y$1:$Y$5' }, values: [v(1), f('"1"'), v(2), f('"2"'), v(true), f('"true"'), v('YES'), v(3)] },
  { rule: { type: 'whole', operator: 'greaterThan', formula1: '0', allowBlank: false }, values: [null, f('""'), v(1)] },
  { rule: { type: 'list', formula1: '"a,b"', allowBlank: false }, values: [null, v('a')] },
  { rule: { type: 'textLength', operator: 'lessThan', formula1: '2', allowBlank: false }, values: [null, f('""'), v('a')] },
]

function book(): BookInput {
  const cells: Record<string, CellInput> = { Y1: { value: 1 }, Y2: { formula: '"2"' }, Y3: { value: true }, Y4: { value: 'Yes' }, Y5: { value: 3.0 } }
  const validations: DataValidation[] = []
  COLUMNS.forEach((col, i) => {
    const letter = String.fromCharCode(65 + i)
    cells[`${letter}1`] = { value: `${col.rule.type} ${i + 1}` }
    col.values.forEach((x, k) => {
      if (x) cells[`${letter}${k + 2}`] = x
    })
    validations.push({ range: `${letter}2:${letter}${col.values.length + 1}`, ...col.rule })
  })
  return { sheets: [{ name: 'S', cells, validations }] }
}

export function buildValidCases(): ValidCases {
  return { books: { values: book() }, cases: [{ id: 'values', book: 'values' }], window: 'A1:U13' }
}
