/**
 * Array formulas of the old kind beyond Excel's answers (`test/excel/spill/cse*.json`, `test/excel/struct/arrays.json`):
 * the input checks, the result's fields, a session's edits of an area, a spill reference to it, and both execution modes.
 */
import { describe, expect, it } from 'vitest'
import { compute, InputError, createSession, type BookInput, type CellInput } from '../src/index.js'

const arr = (formula: string, area: string): CellInput => ({ formula, array: area })
const book = (cells: Record<string, CellInput>): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, ...cells } }] })

describe('array formulas of the old kind', () => {
  it('an area on its own cell, with a formula, holding nothing else, is checked as input', () => {
    expect(() => compute(book({ C1: { value: 1, array: 'C1:C2' } }))).toThrow(/cells\["C1"\]\.array: expected the area of the cell's array formula/)
    expect(() => compute(book({ C2: arr('A1', 'C1:C3') }))).toThrow(InputError)
    expect(() => compute(book({ C1: arr('A1:A3', 'C1:C3'), C2: { value: 5 } }))).toThrow(/cells\["C2"\]: the cell is inside the area of the array formula at C1/)
    expect(() => compute(book({ C1: arr('A1:A3', 'C1:C3'), C3: arr('1', 'C3') }))).toThrow(InputError)
    expect(() => compute(book({ C1: arr('A1:A3', 'C1:C3'), C2: { numberFormat: '0.0' } }))).not.toThrow()
  })

  it('the result lists the area on its cell and the cells of the area as spilled from it', () => {
    const r = compute(book({ C1: arr('A1:A3*2', 'C1:C3'), E1: { formula: 'C1#' }, F1: { formula: 'ROWS(C1#)' } }))
    const c = r.sheets[0].cells
    expect([c.C1.spill, c.C2.spillOf, c.C3.value, c.C3.text]).toEqual(['C1:C3', 'C1', 6, '6'])
    // the spill reference of its cell is its area, which spills (E1:E3) where it can
    expect([c.E1.spill, c.F1.value]).toEqual(['E1:E3', 3])
  })

  it('a session refuses to change part of an area and changes nothing; the whole area, or its formula again, it takes', () => {
    const s = createSession(book({ C1: arr('A1:A3*2', 'C1:C3') }))
    expect(() => s.setCells([{ sheet: 'S', cell: 'C2', input: { value: 9 } }])).toThrow(/C2 is part of the array formula at S!C1/)
    expect(() => s.setCells([{ sheet: 'S', cell: 'C1', input: null }])).toThrow(InputError)
    expect(s.getResult().sheets[0].cells.C2.value).toBe(4)
    expect(s.setCells([{ sheet: 'S', cell: 'C1', input: arr('A1:A3*10', 'C1:C3') }]).changes.map((x) => [x.cell, x.value])).toEqual(expect.arrayContaining([['C2', 20]]))
    s.setCells(['C1', 'C2', 'C3'].map((cell) => ({ sheet: 'S', cell, input: null })))
    expect(s.getResult().sheets[0].cells.C2).toBeUndefined()
    // a new array formula over cells that hold input is refused
    expect(() => s.setCells([{ sheet: 'S', cell: 'A1', input: arr('5', 'A1:A2') }])).toThrow(/A2/)
  })

  it('both execution modes give the same values', () => {
    const b = book({ C1: arr('A1:A3*ROW()', 'C1:D4'), F1: arr('SUM(A1:A3*A1:A3)', 'F1'), G1: { formula: 'SUM(C1:D4)' } })
    const [x, y] = [compute(b, { execution: 'eval', display: false }), compute(b, { execution: 'closures', display: false })]
    expect(x.sheets).toEqual(y.sheets)
    expect([x.sheets[0].cells.D2.value, x.sheets[0].cells.F1.value]).toEqual([4, 14])
  })
})
