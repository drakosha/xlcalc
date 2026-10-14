/**
 * A structure operation that adds a style to the book's table answers the table (`EditResult.styles`), as its undo and
 * its redo do: a copy that spells out a plain style over a styled row, a centred merge that builds the book again.
 */
import { describe, expect, it } from 'vitest'
import { createSession, type BookInput, type StructureOp } from '../src/index.js'

const cases: [string, BookInput, StructureOp][] = [
  ['a copy into a styled row', { sheets: [{ name: 'S', cells: { B1: { value: 1 } }, rows: { 5: { style: 0 } } }], styles: [{ fill: { color: '#FFC000' } }] }, { op: 'copyRange', from: { sheet: 'S', range: 'B1' }, to: { sheet: 'S', range: 'B5' } }],
  ['a centred merge moving a cell', { sheets: [{ name: 'S', cells: { B1: { value: 1 } } }] }, { op: 'mergeCells', sheet: 'S', range: 'A1:B1', center: true }],
  ['a centred merge', { sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] }, { op: 'mergeCells', sheet: 'S', range: 'A1:B1', center: true }],
]

describe('a structure operation that changes the style table', () => {
  for (const [what, book, op] of cases)
    it(`answers it, and so do its undo and redo: ${what}`, () => {
      const s = createSession(book)
      const was = book.styles ?? []
      const r = s.editStructure(op)
      if (!r.ok) throw new Error(r.message)
      const now = s.getBook().styles!
      expect(now.length).toBe(was.length + 1)
      expect(r.styles).toEqual(now)
      expect(s.undo()!.styles ?? []).toEqual(was)
      expect(s.redo()!.styles).toEqual(now)
    })
})
