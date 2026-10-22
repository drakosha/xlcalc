/**
 * Names under structure operations, from the review of 21.10: a sheet rename rewrites a reference to a sheet's own name
 * (`S!loc`), and inserted or deleted lines leave the relative part of a name's mixed reference as an offset, as they
 * leave a reference without `$`, so the cells that use the name keep their values.
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

  describe('inserted and deleted lines leave the relative part of a name mixed reference', () => {
    const mixed = (): BookInput => {
      const cells: BookInput['sheets'][number]['cells'] = { C5: { formula: 'mix' }, E5: { formula: '$A5' }, C14: { formula: 'mixc' }, E14: { formula: 'C$12' } }
      for (let r = 1; r <= 10; r++) cells[`A${r}`] = { value: r * 100 }
      'ABCDEF'.split('').forEach((c, i) => (cells[`${c}12`] = { value: (i + 1) * 7 }))
      return { names: [{ name: 'mix', ref: 'S!$A1' }, { name: 'mixc', ref: 'S!A$12' }, { name: 'rel', ref: 'S!XFD1' }], sheets: [{ name: 'S', cells }] }
    }
    const cases: [StructureOp, string[], [string, string][]][] = [
      [{ op: 'insertRows', sheet: 'S', row: 1 }, ['S!$A1', 'S!A$13', 'S!XFD1'], [['C6', 'E6'], ['C15', 'E15']]],
      [{ op: 'deleteRows', sheet: 'S', row: 1 }, ['S!$A1', 'S!A$11', 'S!XFD1'], [['C4', 'E4'], ['C13', 'E13']]],
      [{ op: 'insertColumns', sheet: 'S', column: 'A' }, ['S!$B1', 'S!A$12', 'S!XFD1'], [['D5', 'F5'], ['D14', 'F14']]],
      [{ op: 'deleteColumns', sheet: 'S', column: 'A' }, ['S!#REF!', 'S!A$12', 'S!XFD1'], [['B5', 'D5'], ['B14', 'D14']]],
      [{ op: 'insertRows', sheet: 'S', row: 3 }, ['S!$A1', 'S!A$13', 'S!XFD1'], [['C6', 'E6'], ['C15', 'E15']]],
    ]

    it.each(cases)('%j', (op, names, pairs) => {
      const b = edited(mixed(), op)
      expect(b.names!.map((n) => n.ref)).toEqual(names)
      for (const [used, control] of pairs) {
        const [a, c] = values(b, 'S', [used, control])
        expect(a).toEqual(c)
      }
    })

    it('an area with a fixed and a relative bound on the shifted axis moves the fixed one only', () => {
      const book: BookInput = { names: [{ name: 'part', ref: 'S!$A$2:A3' }], sheets: [{ name: 'S', cells: {} }] }
      expect(edited(book, { op: 'insertRows', sheet: 'S', row: 1 }).names![0].ref).toBe('S!$A$3:A3')
      expect(edited(book, { op: 'insertRows', sheet: 'S', row: 5, count: 2 }).names![0].ref).toBe('S!$A$2:A3')
      expect(edited(book, { op: 'deleteRows', sheet: 'S', row: 2 }).names![0].ref).toBe('S!#REF!')
    })
  })
})
