/**
 * A copy of a table's cells pasted onto a target that is no multiple of them: one copy at the target's top-left cell, as
 * copyRange lays it, so the target's other cells keep their formulas, a formula in another table's calculated column
 * among them (the set `tables/paste-names-held`, Excel for Mac and Windows).
 */
import { expect, it } from 'vitest'
import { book, sales, table } from '../bench/coverage/tables/cases.js'
import { editStructure, type BookInput } from '../src/index.js'

/** Sales on A1:D5 and Side on F1:G6, whose b is a times 2. */
const SIDE = table('Side', 'F1', ['a', 'b'], [[1, '=[@a]*2'], [3, '=[@a]*2'], [5, '=[@a]*2'], [7, '=[@a]*2'], [9, '=[@a]*2']], { columns: [{}, { formula: 'Side[[#This Row],[a]]*2' }] })

const formulas = (b: BookInput, cells: string[]) => cells.map((a) => b.sheets[0].cells[a]?.formula)

it("a copy onto a target that is no multiple of it leaves the target's other cells as they were", () => {
  const before = book(['S', {}, sales(), SIDE])
  const rows = editStructure(before, { op: 'copyRange', from: { sheet: 'S', range: 'D2:D3' }, to: { sheet: 'S', range: 'G2:G4' } })
  const cols = editStructure(before, { op: 'copyRange', from: { sheet: 'S', range: 'C3:D3' }, to: { sheet: 'S', range: 'E4:G4' } })
  if (!rows.ok || !cols.ok) throw new Error('refused')
  const copied = 'Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]'
  expect(formulas(rows.book, ['G2', 'G3', 'G4'])).toEqual([copied, copied, '[@a]*2'])
  expect(formulas(cols.book, ['F4', 'G4'])).toEqual([copied, '[@a]*2'])
})
