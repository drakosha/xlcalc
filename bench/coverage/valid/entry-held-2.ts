/**
 * Held out after `entry-held` and `lists`: the rules fitted there, on other values, predicted before Excel ran them
 * (`entry-held-2.predicted.json`): literal items that read as numbers otherwise than plainly, spaces and case; formulas
 * referring to a blank cell through a range and through arithmetic; lists through names holding a blank in the middle,
 * a whole column, none; booleans against bounds; blanks where blanks are not ignored; equality to 15 digits.
 */
import type { BookInput, CellInput, DataValidation } from '../../../src/contract.js'
import type { ValidCases } from './cases.js'

type Column = { rule: Omit<DataValidation, 'range'>; values: (CellInput | null)[] }

const v = (value: CellInput['value']): CellInput => ({ value })
const f = (formula: string): CellInput => ({ formula })

const COLUMNS: Column[] = [
  { rule: { type: 'list', formula1: '"1000,50%,2025-01-02, Two ,3"' }, values: [v(1000), f('"1,000"'), v(0.5), v(45659), v('Two'), v('two'), v(' Two '), v(3), f('"3 "')] },
  { rule: { type: 'custom', formula1: 'B2+$Z$1>5' }, values: [v(1), v(10)] },
  { rule: { type: 'custom', formula1: 'SUM($Z$1:$Z$3)>5' }, values: [v(1)] },
  { rule: { type: 'whole', formula1: '$Z$1+1', formula2: '10' }, values: [v(50), v(5)] },
  { rule: { type: 'list', formula1: 'Middle' }, values: [v('x'), v('apple')] },
  { rule: { type: 'list', formula1: 'Column' }, values: [v('x'), v('apple')] },
  { rule: { type: 'list', formula1: 'Full' }, values: [v('x'), v('pear')] },
  { rule: { type: 'list', formula1: 'Middle', allowBlank: false }, values: [v('x'), v('apple')] },
  { rule: { type: 'whole', operator: 'greaterThan', formula1: '5' }, values: [v(true), v(false)] },
  { rule: { type: 'decimal', operator: 'lessThan', formula1: '5' }, values: [v(true), v(false)] },
  { rule: { type: 'date', operator: 'greaterThanOrEqual', formula1: '1' }, values: [v(true), v('x')] },
  { rule: { type: 'textLength', operator: 'greaterThan', formula1: '0', allowBlank: false }, values: [null, v('a')] },
  { rule: { type: 'whole', formula1: '-1', formula2: '1', allowBlank: false }, values: [null, v(0)] },
  { rule: { type: 'custom', formula1: 'TRUE', allowBlank: false }, values: [null, v(1)] },
  { rule: { type: 'decimal', operator: 'equal', formula1: '1/3' }, values: [v(0.333333333333333), v(0.3333333333333333), v(0.33333333333333)] },
  { rule: { type: 'decimal', operator: 'notEqual', formula1: '0.3' }, values: [f('0.1+0.2'), v(0.31)] },
]

function book(): BookInput {
  const cells: Record<string, CellInput> = { U1: v('apple'), U3: v('pear'), U5: v('plum'), W1: v('pear'), W2: v('fig') }
  const validations: DataValidation[] = []
  COLUMNS.forEach((col, i) => {
    const letter = String.fromCharCode(65 + i)
    cells[`${letter}1`] = { value: `${col.rule.type} ${i + 1}` }
    col.values.forEach((x, k) => {
      if (x) cells[`${letter}${k + 2}`] = x
    })
    validations.push({ range: `${letter}2:${letter}${col.values.length + 1}`, ...col.rule })
  })
  return {
    sheets: [{ name: 'S', cells, validations }],
    names: [
      { name: 'Middle', ref: "'S'!$U$1:$U$3" },
      { name: 'Column', ref: "'S'!$U:$U" },
      { name: 'Full', ref: "'S'!$W$1:$W$2" },
    ],
  }
}

export function buildValidCases(): ValidCases {
  return { books: { values: book() }, cases: [{ id: 'values', book: 'values' }], window: 'A1:P10' }
}
