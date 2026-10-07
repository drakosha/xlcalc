/**
 * Copying and moving spilled cells and anchors within their sheet: a spilled cell alone, part of a spill range,
 * the whole range with its anchor, the anchor with part of its range. Dynamic array formulas, so the anchors
 * spill. Moving an anchor to another sheet crashed Excel 16.113 twice and is not a case; the moves of spilled
 * cells without their anchor come last.
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import type { StructCase, StructCases } from './cases.js'

const f = (formula: string): CellInput => ({ formula })

/** S: E2 spills E2:E4, B1 spills B1:C1; G1 sums E2#, H1 reads E3. */
function book(): BookInput {
  return { sheets: [{ name: 'S', cells: { A1: { value: 1 }, B1: f('SEQUENCE(1,2)'), E2: f('SEQUENCE(3)'), G1: f('SUM(E2#)'), H1: f('E3+1') } }] }
}

export function buildStructCases(): StructCases {
  const cases: StructCase[] = []
  const copy = (range: string, to: string) => cases.push({ id: `copy-${range}-${to}`, book: 'spill', op: { op: 'copyRange', from: { sheet: 'S', range }, to: { sheet: 'S', range: to } } })
  const move = (range: string, to: string) => cases.push({ id: `move-${range}-${to}`, book: 'spill', op: { op: 'moveRange', from: { sheet: 'S', range }, to: { sheet: 'S', cell: to } } })
  copy('E3', 'J3')
  copy('E3:E4', 'J3')
  copy('C1', 'C10')
  copy('E2:E4', 'J2')
  copy('E2:E3', 'J2')
  copy('E2', 'J2')
  move('E2:E4', 'J2')
  move('E2:E3', 'J2')
  move('E3', 'J3')
  move('E3:E4', 'J3')
  return { books: { spill: book() }, cases, dynamic: true }
}
