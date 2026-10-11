/** A rectangle's cells: their count, listed row by row, or past a count those of a sheet's cells within it (`cellsWithin`). */
import { describe, expect, it } from 'vitest'
import { cellCount, cellsWithin, parseRange, rectCells } from '../src/structure/area.js'

describe("a rectangle's cells", () => {
  it('row by row, and their count', () => {
    const x = parseRange('B2:C3')!
    expect([...rectCells(x)]).toEqual(['B2', 'C2', 'B3', 'C3'])
    expect(cellCount(x)).toBe(4)
    expect(cellCount(parseRange('A:A')!)).toBe(1_048_576)
  })

  it("up to the count listed, whether the sheet has them or not; past it the sheet's cells within, in the sheet's order", () => {
    const sheet = ['C3', 'A1', 'B2', 'Z9']
    const x = parseRange('A1:B2')!
    expect([...cellsWithin(x, () => sheet, 4)]).toEqual(['A1', 'B1', 'A2', 'B2'])
    expect([...cellsWithin(x, () => sheet, 3)]).toEqual(['A1', 'B2'])
    expect([...cellsWithin(parseRange('B:C')!, () => sheet)]).toEqual(['C3', 'B2'])
  })
})
