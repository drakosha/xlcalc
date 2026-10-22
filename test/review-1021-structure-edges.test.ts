/**
 * Reproductions from the 21.10 review, structure at its edges: a fill over dates that runs past the date system writes
 * the serial numbers, as a fill over numbers in a date format does. `limits.writes` bounds a fill's cells all together,
 * in a synchronous session, in the same thread and in a thread. A copy or a fill keeps the spelling of the references
 * it does not change. Inserting lines refuses to push a legacy array's area off the sheet, and a shift that leaves no
 * row, column or merged area leaves no empty field for them.
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

  describe('a copy keeps the spelling of a reference it does not change', () => {
    const book = (): BookInput => ({ sheets: [{ name: 'S', cells: { B1: { formula: 'sum( a1 , $a$1 ) + $A$1 + s!$a$1 + S!a1' }, C1: { formula: '$a$1*2 + a:a + sum(1:1)' } } }, { name: 'T', cells: {} }] })

    it.each([
      ['copyRange', { op: 'copyRange', from: { sheet: 'S', range: 'B1:C1' }, to: { sheet: 'S', range: 'B2:C2' } }],
      ['fillRange', { op: 'fillRange', sheet: 'S', range: 'B1:C1', to: 'B1:C2' }],
    ] as [string, StructureOp][])('%s', (_, op) => {
      const cells = edited(book(), op, execution)
      expect(cells.B2.formula).toBe('sum( A2 , $a$1 ) + $A$1 + s!$a$1 + S!A2')
      expect(cells.C2.formula).toBe('$a$1*2 + a:a + sum(2:2)')
    })

    it('to another sheet a copy changes nothing; a move spells out the sheet as Excel writes it', () => {
      const r = editStructure(book(), { op: 'copyRange', from: { sheet: 'S', range: 'C1' }, to: { sheet: 'T', range: 'C1' } })
      expect(r.ok && r.book.sheets[1].cells.C1.formula).toBe('$a$1*2 + a:a + sum(1:1)')
      const m = editStructure(book(), { op: 'moveRange', from: { sheet: 'S', range: 'C1' }, to: { sheet: 'T', cell: 'C1' } })
      expect(m.ok && m.book.sheets[1].cells.C1.formula).toBe("'S'!$A$1*2 + 'S'!A:A + sum('S'!1:1)")
    })
  })

  describe("inserting lines refuses to push a legacy array's area off the sheet", () => {
    it.each([
      ['rows', { A1048575: { formula: '{1;2}', array: 'A1048575:A1048576' } }, { op: 'insertRows', sheet: 'S', row: 1 }, 'A1048575'],
      ['columns', { XFC1: { formula: '{1,2}', array: 'XFC1:XFD1' } }, { op: 'insertColumns', sheet: 'S', column: 'A' }, 'XFC1'],
      ['two rows', { A1048573: { formula: '{1;2;3}', array: 'A1048573:A1048575' } }, { op: 'insertRows', sheet: 'S', row: 1, count: 2 }, 'A1048573'],
    ] as [string, BookInput['sheets'][number]['cells'], StructureOp, string][])('%s', (_, cells, op, top) => {
      const r = editStructure(one(cells), op)
      expect(r).toMatchObject({ ok: false, code: 'off_sheet' })
      expect(!r.ok && r.message).toContain(`S!${top}`)
    })

    it('an area that stays on the sheet moves', () => {
      const cells = edited(one({ A1048574: { formula: '{1;2}', array: 'A1048574:A1048575' } }), { op: 'insertRows', sheet: 'S', row: 1 }, execution)
      expect(cells).toEqual({ A1048575: { formula: '{1;2}', array: 'A1048575:A1048576' } })
    })
  })

  describe('a shift that leaves no rows, columns or merged areas drops their fields', () => {
    it.each([
      ['a row pushed off', { rows: { 1048576: { outlineLevel: 1 } } }, { op: 'insertRows', sheet: 'S', row: 1 }],
      ['the only row deleted', { rows: { 5: { height: 20 } } }, { op: 'deleteRows', sheet: 'S', row: 5 }],
      ['columns and a merged area deleted', { columns: { XFD: { width: 20 } }, merges: ['XFC1:XFD1'] }, { op: 'deleteColumns', sheet: 'S', column: 'XFC', count: 2 }],
    ] as [string, Partial<BookInput['sheets'][number]>, StructureOp][])('%s', (_, layout, op) => {
      const r = editStructure({ sheets: [{ name: 'S', cells: { B2: { value: 1 } }, ...layout }] }, op)
      expect(r.ok && Object.keys(r.book.sheets[0])).toEqual(['name', 'cells'])
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
