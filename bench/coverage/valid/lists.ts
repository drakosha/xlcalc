/**
 * Held out: lists of cells, the items a dropdown offers, as `Validation.Value` takes a value for one of them, predicted
 * before Excel ran them (`lists.predicted.json`): a list through a defined name (global and local to its sheet), a range
 * of another sheet, a spilled array (`W1#`), a formula giving a reference, a whole column, a blank among the items, and
 * literal lists with `;` and spaces in them.
 */
import type { BookInput, CellInput, DataValidation } from '../../../src/contract.js'
import type { ValidCases } from './cases.js'

const v = (value: CellInput['value']): CellInput => ({ value })
const f = (formula: string): CellInput => ({ formula })

/** Each list and the values tried against it, by column from A. */
const COLUMNS: { formula1: string; allowBlank?: boolean; values: CellInput[] }[] = [
  { formula1: 'Fruits', values: [v('apple'), v('PEAR'), v(3), v('kiwi'), f('"3"')] },
  { formula1: 'Local', values: [v('pear'), v('apple'), v(3)] },
  { formula1: "'L'!$A$1:$A$3", values: [v('Y'), v('w'), v('x ')] },
  { formula1: '$W$1#', values: [v(11), v(13), f('"11"')] },
  { formula1: 'INDEX($U$1:$U$4,0,1)', values: [v('apple'), v('kiwi')] },
  { formula1: '$U:$U', values: [v('apple'), v(3), v('kiwi')] },
  { formula1: '$U$1:$U$5', values: [f('""'), v('Pear')] },
  { formula1: '$U$1:$U$5', allowBlank: false, values: [f('""'), v('apple')] },
  { formula1: '"a;b;c"', values: [v('b'), v('a;b;c'), v('a')] },
  { formula1: '"a, b,c "', values: [v('b'), v(' b'), v('c '), v('c')] },
]

function book(): BookInput {
  const cells: Record<string, CellInput> = { U1: v('apple'), U2: v('Pear'), U3: v(3), U5: v('plum'), W1: f('SEQUENCE(3,1,10)') }
  const validations: DataValidation[] = []
  COLUMNS.forEach((col, i) => {
    const letter = String.fromCharCode(65 + i)
    cells[`${letter}1`] = { value: col.formula1 }
    col.values.forEach((x, k) => (cells[`${letter}${k + 2}`] = x))
    validations.push({ range: `${letter}2:${letter}${col.values.length + 1}`, type: 'list', formula1: col.formula1, ...(col.allowBlank === false ? { allowBlank: false } : {}) })
  })
  return {
    sheets: [
      { name: 'S', cells, validations },
      { name: 'L', cells: { A1: v('x'), A2: v('y'), A3: v('z') } },
    ],
    names: [
      { name: 'Fruits', ref: "'S'!$U$1:$U$4" },
      { name: 'Local', ref: "'S'!$U$2:$U$3", scope: 'S' },
    ],
  }
}

export function buildValidCases(): ValidCases {
  return { books: { lists: book() }, cases: [{ id: 'lists', book: 'lists' }], window: 'A1:J7' }
}
