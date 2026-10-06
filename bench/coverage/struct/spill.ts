/**
 * Spill references (`E2#`) under structure operations: they move as references to their anchor cell, in
 * formulas of the anchor's sheet and of another sheet and in a name; deleting the anchor or its sheet loses
 * them. The formulas are dynamic array formulas, so the anchors spill as in Excel 365.
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import type { StructCase, StructCases } from './cases.js'

const f = (formula: string): CellInput => ({ formula })

/** S: E2 spills E2:E4, B1 spills B1:C1, readers of both; T: readers from another sheet; a name over E2#. */
function spillBook(): BookInput {
  return {
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { value: 1 },
          B1: f('SEQUENCE(1,2)'),
          E2: f('SEQUENCE(3)'),
          G1: f('SUM(E2#)'),
          G2: f('E2#*2'),
          G6: f('ROWS($E$2#)'),
          G7: f('SUM(sp)'),
          D8: f('B1#'),
          H1: f('E3+1'),
        },
      },
      { name: 'T', cells: { A1: f('SUM(S!E2#)'), A2: f('S!$E$2#'), B1: f('COLUMNS(S!B1#)') } },
    ],
    names: [{ name: 'sp', ref: "'S'!$E$2#" }],
  }
}

/** The book of the spill cases. */
export { spillBook }

export function buildStructCases(): StructCases {
  const cases: StructCase[] = []
  const add = (id: string, op: StructCase['op']) => cases.push({ id, book: 'spill', op })
  add('insert-rows-1', { op: 'insertRows', sheet: 'S', row: 1 })
  add('insert-rows-3', { op: 'insertRows', sheet: 'S', row: 3 })
  add('delete-rows-2', { op: 'deleteRows', sheet: 'S', row: 2 })
  add('delete-rows-3', { op: 'deleteRows', sheet: 'S', row: 3 })
  add('insert-columns-A', { op: 'insertColumns', sheet: 'S', column: 'A' })
  add('delete-columns-E', { op: 'deleteColumns', sheet: 'S', column: 'E' })
  add('delete-columns-B', { op: 'deleteColumns', sheet: 'S', column: 'B' })
  add('move-E2-J2', { op: 'moveRange', from: { sheet: 'S', range: 'E2' }, to: { sheet: 'S', cell: 'J2' } })
  // cutting E2 to another sheet ('T'!C5) crashed Excel 16.113 twice (06.10): not a case
  return { books: { spill: spillBook() }, cases, dynamic: true }
}
