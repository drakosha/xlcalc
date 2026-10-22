/**
 * Held out: the cell a rule's relative formula is relative to when the rule's ranges do not start at their corner. A
 * custom rule `$A3>2` over H3:H6 and J1:J3: made by Excel with the ranges selected in that order (H3 active), and
 * written by `writeXlsx` with its ranges in that order and in the order Excel saves (`J1:J3 H3:H6`). Column A holds its
 * row's number and the ruled cells 0, so each cell's verdict and formula show the row of A it reads. Predicted before
 * Excel ran it (`anchor-held.predicted.json`); the file Excel saved shows the `sqref` and `formula1` it writes.
 */
import type { BookInput, CellInput, DataValidation } from '../../../src/contract.js'
import type { ValidCases } from './cases.js'

const RULE: Omit<DataValidation, 'range'> = { type: 'custom', formula1: '$A3>2' }

/** S: 1 to 6 in A1:A6, 0 in H1:J6; the rule over `range` when given. */
function book(range?: string): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= 6; r++) {
    cells[`A${r}`] = { value: r }
    for (const c of 'HIJ') cells[`${c}${r}`] = { value: 0 }
  }
  return { sheets: [{ name: 'S', cells, ...(range ? { validations: [{ range, ...RULE }] } : {}) }] }
}

export function buildValidCases(): ValidCases {
  return {
    books: { plain: book(), written: book('H3:H6 J1:J3'), saved: book('J1:J3 H3:H6') },
    cases: [
      { id: 'made', book: 'plain', make: [{ sheet: 'S', range: 'H3:H6 J1:J3', rule: RULE }] },
      { id: 'written', book: 'written' },
      { id: 'written-corner-last', book: 'saved' },
    ],
  }
}
