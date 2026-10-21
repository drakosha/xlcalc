/**
 * Names under structure operations, from the review of 21.10: a sheet rename rewrites a reference to a sheet's own name
 * (`S!loc`).
 */
import { describe, expect, it } from 'vitest'
import { compute, createSession, editStructure, type BookInput, type CellValue, type StructureOp } from '../src/index.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

/** The book after the operation, which must succeed. */
function edited(book: BookInput, op: StructureOp): BookInput {
  const r = editStructure(book, op)
  if (!r.ok) throw new Error(r.code)
  return r.book
}

const qualified = (): BookInput => ({
  names: [{ name: 'loc', ref: 'S!$A$1', scope: 'S' }, { name: 'g', ref: 'S!loc+1' }],
  sheets: [
    { name: 'S', cells: { A1: { value: 42 } } },
    { name: 'T', cells: { A1: { formula: 'S!loc*2' }, A2: { formula: "SUM('S'!loc)" }, A3: { formula: 'g' } } },
  ],
})

describe.each(['closures', 'eval'] as const)('%s', (execution) => {
  const values = (book: BookInput, sheet: string, cells: string[]): CellValue[] => {
    const r = compute(book, { execution, clock })
    return cells.map((c) => r.sheets.find((x) => x.name === sheet)!.cells[c]?.value ?? null)
  }

  describe('a sheet rename rewrites references to the sheet own names', () => {
    it('in cells and in names, and the values stay', () => {
      expect(values(qualified(), 'T', ['A1', 'A2', 'A3'])).toEqual([84, 42, 43])
      for (const name of ['U', 'My Sheet']) {
        const b = edited(qualified(), { op: 'renameSheet', sheet: 'S', name })
        const q = `'${name}'`
        expect(b.sheets[1].cells).toEqual({ A1: { formula: `${q}!loc*2` }, A2: { formula: `SUM(${q}!loc)` }, A3: { formula: 'g' } })
        expect(b.names).toEqual([{ name: 'loc', ref: `'${name}'!$A$1`, scope: name }, { name: 'g', ref: `${q}!loc+1` }])
        expect(values(b, 'T', ['A1', 'A2', 'A3'])).toEqual([84, 42, 43])
      }
    })

    it('in a session', () => {
      const s = createSession(qualified(), { execution, clock })
      expect(s.editStructure({ op: 'renameSheet', sheet: 'S', name: 'U' }).ok).toBe(true)
      expect(s.getCells('T', 'A1:A3')).toMatchObject({ A1: { value: 84 }, A2: { value: 42 }, A3: { value: 43 } })
      expect(s.getHash()).toBe(compute(s.getBook(), { execution, clock }).hash)
    })
  })
})
