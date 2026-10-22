/**
 * Reproductions from the 21.10 review, structure at its edges: a fill over dates that runs past the date system writes
 * the serial numbers, as a fill over numbers in a date format does. `limits.writes` bounds a fill's cells all together,
 * in a synchronous session, in the same thread and in a thread.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { CalcError, compute, editStructure, type BookInput, type StructureOp } from '../src/index.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

afterAll(closeCalcs)

const one = (cells: BookInput['sheets'][number]['cells'], dateSystem?: 1904): BookInput => ({ ...(dateSystem ? { dateSystem } : {}), sheets: [{ name: 'S', cells }] })

/** The cells of the book an edit gives, after checking that it computes. */
function edited(book: BookInput, op: StructureOp, execution: 'closures' | 'eval') {
  const r = editStructure(book, op)
  if (!r.ok) throw new Error(`refused: ${r.code}`)
  compute(r.book, { clock: CLOCK, execution })
  return r.book.sheets[0].cells
}

describe.each(['closures', 'eval'] as const)('%s', (execution) => {
  describe('a fill over dates past the date system', () => {
    it('up from 1900-01-02: the serials below day 0 as numbers, the dates as dates', () => {
      const cells = edited(one({ A5: { value: { date: '1900-01-02' } }, A6: { value: { date: '1900-01-03' } } }), { op: 'fillRange', sheet: 'S', range: 'A5:A6', to: 'A1:A6' }, execution)
      expect([1, 2, 3, 4].map((r) => cells[`A${r}`].value)).toEqual([-2, -1, { date: '1900-01-00' }, { date: '1900-01-01' }])
    })

    it('down past 9999-12-31, keeping the number format', () => {
      const cells = edited(one({ A1: { value: { date: '9999-12-30' }, numberFormat: 'yyyy-mm-dd' }, A2: { value: { date: '9999-12-31' }, numberFormat: 'yyyy-mm-dd' } }), { op: 'fillRange', sheet: 'S', range: 'A1:A2', to: 'A1:A4' }, execution)
      expect([cells.A3, cells.A4]).toEqual([{ value: 2958466, numberFormat: 'yyyy-mm-dd' }, { value: 2958467, numberFormat: 'yyyy-mm-dd' }])
    })

    it('up by months, and in the 1904 system', () => {
      const months = edited(one({ A5: { value: { date: '1900-02-01' } }, A6: { value: { date: '1900-03-01' } } }), { op: 'fillRange', sheet: 'S', range: 'A5:A6', to: 'A1:A6' }, execution)
      expect(months.A4.value).toEqual({ date: '1900-01-01' })
      expect(typeof months.A3.value).toBe('number')
      const d1904 = edited(one({ A2: { value: { date: '1904-01-02' } } }, 1904), { op: 'fillRange', sheet: 'S', range: 'A2', to: 'A1:A2', type: 'series' }, execution)
      expect(d1904.A1.value).toEqual({ date: '1904-01-01' })
      const below = edited(one({ A3: { value: { date: '1904-01-02' } } }, 1904), { op: 'fillRange', sheet: 'S', range: 'A3', to: 'A1:A3', type: 'series' }, execution)
      expect(below.A1.value).toBe(-1)
    })
  })

  describe.each(KINDS)('limits.writes, %s', (kind) => {
    const open = openerFor(kind)
    const ten = () => one(Object.fromEntries(Array.from({ length: 10 }, (_, i) => [`A${i + 1}`, { value: i + 1 }])))

    it.each(['copy', 'default'] as const)('a fill of type %s counts its cells together', async (type) => {
      const s = await open(ten(), { clock: CLOCK, execution, limits: { writes: 15 } })
      // 19 cells, laid as a copy of the 10 and one of 9: each under the limit, together over it
      const e = await s.editStructure({ op: 'fillRange', sheet: 'S', range: 'A1:A10', to: 'A1:A29', type }).then(() => null, (x: unknown) => x)
      expect(e).toBeInstanceOf(CalcError)
      expect(e).toMatchObject({ code: 'limit', limit: 'writes', size: 19, max: 15 })
      expect(Object.keys((await s.getBook()).sheets[0].cells)).toHaveLength(10)
      expect(await s.editStructure({ op: 'fillRange', sheet: 'S', range: 'A1:A10', to: 'A1:A25', type })).toMatchObject({ ok: true })
      await s.close()
    })
  })
})
