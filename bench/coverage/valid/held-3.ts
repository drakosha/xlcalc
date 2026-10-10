/**
 * Held out after `struct-held-2` and `entry-held-2`: the rules fitted there, predicted before Excel ran them
 * (`held-3.predicted.json`): blanks where blanks are not ignored against every type, literal items in other number
 * spellings, rules with absolute and relative references moved and copied to another sheet.
 */
import type { BookInput, CellInput, DataValidation } from '../../../src/contract.js'
import type { ValidCases } from './cases.js'

type Column = { rule: Omit<DataValidation, 'range'>; values: (CellInput | null)[] }

const v = (value: CellInput['value']): CellInput => ({ value })

const COLUMNS: Column[] = [
  { rule: { type: 'decimal', operator: 'lessThan', formula1: '1', allowBlank: false }, values: [null, v(2)] },
  { rule: { type: 'date', operator: 'greaterThan', formula1: '0', allowBlank: false }, values: [null, v(1)] },
  { rule: { type: 'time', operator: 'notEqual', formula1: '0', allowBlank: false }, values: [null, v(0.5)] },
  { rule: { type: 'textLength', operator: 'equal', formula1: '0', allowBlank: false }, values: [null, v('a')] },
  { rule: { type: 'custom', formula1: 'ISBLANK(E2)', allowBlank: false }, values: [null, v(1)] },
  { rule: { type: 'list', formula1: '$Y$1:$Y$3', allowBlank: false }, values: [null, v('y')] },
  { rule: { type: 'list', formula1: '"1.5,1e3,+2,.5,0x10"' }, values: [v(1.5), v(1000), v(2), v(0.5), v(16), v('+2')] },
]

function values(): BookInput {
  const cells: Record<string, CellInput> = { Y1: v('y'), Y3: v('z') }
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

/** S: rules with absolute and relative references over B2:C3; T: an empty sheet to move and copy them to. */
function moves(): BookInput {
  const cells: Record<string, CellInput> = { A1: v(1), A2: v(2), A3: v(3), E1: v(10) }
  return {
    sheets: [
      {
        name: 'S',
        cells,
        validations: [
          { range: 'B2:B3', type: 'whole', operator: 'lessThan', formula1: '$E$1' },
          { range: 'C2:C3', type: 'custom', formula1: 'C2>$A$1+A2' },
        ],
      },
      { name: 'T', cells: { A1: v(5) } },
    ],
  }
}

export function buildValidCases(): ValidCases {
  return {
    books: { values: values(), moves: moves() },
    cases: [
      { id: 'values', book: 'values' },
      { id: 'move-to-t', book: 'moves', op: { op: 'moveRange', from: { sheet: 'S', range: 'B2:C3' }, to: { sheet: 'T', cell: 'C4' } } },
      { id: 'copy-to-t', book: 'moves', op: { op: 'copyRange', from: { sheet: 'S', range: 'B2:C3' }, to: { sheet: 'T', range: 'C4' } } },
    ],
    window: 'A1:G8',
  }
}
