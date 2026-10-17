/**
 * `getValues` gives a rectangle's values as rows, each the value `readCells` gives in its field `value` and null where
 * a cell holds none: over the varied books of `read-answers.ts`, a large spill's unstored cells, after edits; a whole
 * column with `maxCells`, refused without; arrays of each call's own; in every kind of session and both executions.
 */
import { expect, it } from 'vitest'
import { CalcError, InputError, type AsyncSession, type CellValue } from '../src/index.js'
import { BOOKS, CLOCK } from './read-answers.js'
import { inEverySession } from './recorded.js'

/** The rows `readCells` gives for a rectangle of the sheet (its first `rows` rows and `cols` columns from A1), null where it lists no cell. */
async function viaReadCells(s: AsyncSession, sheet: string, range: string, rows: number, cols: number): Promise<CellValue[][]> {
  const out: CellValue[][] = Array.from({ length: rows }, () => Array<CellValue>(cols).fill(null))
  const { row: r0, col: c0 } = corner(range)
  for (const c of (await s.readCells([{ sheet, range }], { fields: ['value'], maxCells: rows * cols })).cells) {
    const { row, col } = corner(c.cell)
    out[row - r0][col - c0] = c.value ?? null
  }
  return out
}

/** A range's top-left cell, zero-based. */
function corner(range: string): { row: number; col: number } {
  const m = /^([A-Z]+)(\d+)/.exec(range)!
  let col = 0
  for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64
  return { row: Number(m[2]) - 1, col: col - 1 }
}

inEverySession('getValues', (execution, open) => {
  it('gives the values readCells gives, null where a cell holds none, before and after edits', async () => {
    for (const book of Object.values(BOOKS)) {
      const s = await open(book, { clock: CLOCK, execution })
      for (const sheet of book.sheets.map((x) => x.name)) {
        expect(await s.getValues(sheet, 'A1:L12'), sheet).toEqual(await viaReadCells(s, sheet, 'A1:L12', 12, 12))
        expect(await s.getValues(sheet, 'B2:D5'), sheet).toEqual(await viaReadCells(s, sheet, 'B2:D5', 4, 3))
      }
      const [first] = book.sheets
      await s.setCells([{ sheet: first.name, cell: 'A1', input: { value: 7 } }, { sheet: first.name, cell: 'C2', input: { formula: 'A1*3' } }])
      const after = await s.getValues(first.name, 'A1:L12')
      expect(after).toEqual(await viaReadCells(s, first.name, 'A1:L12', 12, 12))
      expect(after[0][0]).toBe(7)
      expect(after[1][2]).toBe(21)
      await s.close()
    }
  })

  it("a large spill's unstored cells, a whole column with maxCells and refused without, each call's own arrays", async () => {
    const s = await open({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { value: 2 }, B1: { formula: 'A:A*2' }, C1: { value: 0.1 }, C2: { formula: 'C1*3' }, D1: { formula: '1/0' } } }] }, { clock: CLOCK, execution })
    expect(await s.getValues('S', 'A1:D3')).toEqual([[1, 2, 0.1, { error: 'DIV0' }], [2, 4, 0.3, null], [null, 0, null, null]])
    expect(await s.getValues('S', 'B1000000')).toEqual([[0]])
    await expect(s.getValues('S', 'B:B')).rejects.toBeInstanceOf(CalcError)
    await expect(s.getValues('S', 'B:B')).rejects.toThrow('the range has 1048576 cells, over the limit of 1000000')
    const column = await s.getValues('S', 'B:B', { maxCells: 1048576 })
    expect(column.length).toBe(1048576)
    expect([column[0], column[1], column[1048575]]).toEqual([[2], [4], [0]])
    const rows = await s.getValues('S', '1:1', { maxCells: 16384 })
    expect(rows[0].slice(0, 5)).toEqual([1, 2, 0.1, { error: 'DIV0' }, null])
    const mine = await s.getValues('S', 'A1:A2')
    mine[0][0] = 99
    expect(await s.getValues('S', 'A1:A2')).toEqual([[1], [2]])
    await s.close()
  })

  it('refuses an unknown sheet, a range that is none and options it does not take', async () => {
    const s = await open({ sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] }, { clock: CLOCK, execution })
    await expect(s.getValues('T', 'A1')).rejects.toBeInstanceOf(InputError)
    await expect(s.getValues('S', 'A0')).rejects.toThrow('range')
    await expect(s.getValues('S', 'A1', { maxCells: 0 })).rejects.toThrow('opts.maxCells')
    await expect(s.getValues('S', 'A1', { rows: [1] } as object)).rejects.toThrow('rows')
    expect(await s.getValues('S', 'A1:B1', { maxCells: 2 })).toEqual([[1, null]])
    await expect(s.getValues('S', 'A1:B1', { maxCells: 1 })).rejects.toBeInstanceOf(CalcError)
    await s.close()
  })
})
