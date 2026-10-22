/**
 * Array formulas of the old kind, from the review of 21.10: a copy of part of an area pastes each copied cell with its
 * own note, link, meta and looks, not the area's cell's (fix-u301). `editStructure` and the three kinds of session
 * give the same book, undo and redo it; the session's `getBook()` passes checkBook and its hash is a full computation's.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, editStructure, type BookInput, type Execution, type StructureOp } from '../src/index.js'
import { checkBook } from '../src/input.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const EXECUTIONS: Execution[] = ['closures', 'eval']

afterAll(closeCalcs)

/** B1:B3 an array formula whose cell carries a note, a link, meta and looks; B2 its own note, link and meta; B3 no entry. */
const book = (): BookInput => ({
  styles: [{}, { font: { bold: true } }],
  sheets: [{
    name: 'S',
    cells: {
      A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 },
      B1: { formula: 'A1:A3*2', array: 'B1:B3', note: { text: 'top only' }, link: { url: 'https://example.com' }, meta: { id: 7 }, style: 1, numberFormat: '0.00' },
      B2: { note: { text: 'second' }, link: { location: 'S!A1' }, meta: { id: 8 } },
    },
  }],
})

const copy = (range: string, withMeta: boolean): StructureOp => ({ op: 'copyRange', from: { sheet: 'S', range }, to: { sheet: 'S', range: 'D1' }, withMeta })

const CASES: [string, StructureOp, BookInput['sheets'][number]['cells']][] = [
  ['B1:B2 with meta', copy('B1:B2', true), {
    D1: { formula: 'C1:C3*2', array: 'D1', note: { text: 'top only' }, link: { url: 'https://example.com' }, meta: { id: 7 }, style: 1, numberFormat: '0.00' },
    D2: { formula: 'C2:C4*2', array: 'D2', note: { text: 'second' }, link: { location: 'S!A1' }, meta: { id: 8 } },
  }],
  ['B2:B3 without meta', copy('B2:B3', false), {
    D1: { formula: 'C1:C3*2', array: 'D1', note: { text: 'second' }, link: { location: 'S!A1' } },
    D2: { formula: 'C2:C4*2', array: 'D2' },
  }],
]

const pasted = (b: BookInput) => ({ D1: b.sheets[0].cells.D1, D2: b.sheets[0].cells.D2 })

describe('a copy of part of an array formula of the old kind (fix-u301)', () => {
  it.each(CASES)('%s: each pasted cell keeps its own note, link, meta and looks', (_, op, cells) => {
    const r = editStructure(book(), op)
    if (!r.ok) throw new Error(r.message)
    expect(pasted(r.book)).toEqual(cells)
    expect(() => checkBook(r.book)).not.toThrow()
  })
})

describe.each(EXECUTIONS)('a copy of part of an array formula of the old kind in a session (fix-u301), %s', (execution) => {
  describe.each(KINDS)('in %s', (kind) => {
    const open = openerFor(kind)
    it.each(CASES)('%s: the book editStructure gives, undone and redone', async (_, op, cells) => {
      const s = await open(book(), { execution, clock })
      const before = structuredClone(await s.getBook())
      expect(await s.editStructure(op)).toMatchObject({ ok: true })
      const after = structuredClone(await s.getBook())
      expect(pasted(after)).toEqual(cells)
      expect(() => checkBook(after)).not.toThrow()
      expect(await s.getHash()).toBe(compute(after, { execution, clock }).hash)
      await s.undo()
      expect(await s.getBook()).toEqual(before)
      await s.redo()
      expect(await s.getBook()).toEqual(after)
      await s.close()
    })
  })
})
