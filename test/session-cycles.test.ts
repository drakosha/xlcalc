/**
 * Circular references without iterative calculation in a session (`CellInput.held`, `docs/cycles.md`), by the
 * scenarios measured in Excel for Windows (the workbook set `cycles`): a book opened with the values its file stores,
 * a precedent edited, a formula entered into a cycle's way, the cycle broken by a value and closed again, undo and
 * redo, previews, typed text, looks, a structure operation and a copy. Every kind of session in both execution modes;
 * after each step the session's hash is that of a full computation of its book.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { compute, type AsyncSession, type BookInput, type CachedValues, type CellInput, type Execution } from '../src/index.js'
import { closeCalcs, KINDS, keptPool, opener } from './session-kinds.js'
import { createInlineCalc } from '../src/index.js'

const clock = { now: '2026-10-07T12:00:00Z', seed: 1 }
const LOOP: Record<string, CellInput> = { A1: { formula: 'B1/2+1' }, B1: { formula: 'A1' }, C1: { formula: 'A1*10' }, K1: { value: 3 }, E1: { formula: 'K1*2' } }
const book = (cells = LOOP): BookInput => ({ sheets: [{ name: 'S', cells }] })
/** The values the file of `stored` stores (`cycles:stored`): 2, 2 and 0, and a stale 0 for E1, which reads no cycle. */
const STORED: CachedValues = { sheets: [{ name: 'S', cells: { A1: { value: 2 }, B1: { value: 2 }, C1: { value: 0 }, E1: { value: 0 } } }] }
afterAll(closeCalcs)

describe.each(['closures', 'eval'] as Execution[])('circular references in a session, %s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    let open: Awaited<ReturnType<typeof opener>>
    beforeAll(async () => {
      open = await opener(kind)
    })
    /** The values of cells `refs` of sheet S, after checking that the session gives what a full computation of its book gives. */
    const at = async (s: AsyncSession, refs = 'A1:E1') => {
      const b = await s.getBook()
      expect(await s.getHash()).toBe(compute(b, { clock: s.clock, execution }).hash)
      const cells = await s.getCells('S', refs)
      return Object.fromEntries(Object.entries(cells).map(([k, c]) => [k, c.value]))
    }
    const held = async (s: AsyncSession) => Object.fromEntries(Object.entries((await s.getBook()).sheets[0].cells).flatMap(([k, c]) => (c.held === undefined ? [] : [[k, c.held]])))
    const set = (s: AsyncSession, cell: string, input: CellInput | null) => s.setCells([{ sheet: 'S', cell, input }])

    it('opens with the stored values; a precedent edited and a formula entered; broken by a value, closed again; undo and redo', async () => {
      const s = await open(book({ ...LOOP, A1: { formula: 'B1/2+K1' } }), { clock, execution, cached: STORED })
      try {
        expect(await at(s)).toEqual({ A1: 2, B1: 2, C1: 0, E1: 6 })
        expect(await held(s)).toEqual({ A1: 2, B1: 2 })
        // a precedent of the cycle: nothing in its way computes
        await set(s, 'K1', { value: 5 })
        expect(await at(s)).toEqual({ A1: 2, B1: 2, C1: 0, E1: 10 })
        // a formula entered in its way holds 0, the others what they hold
        await set(s, 'B1', { formula: 'A1*3' })
        expect(await at(s)).toEqual({ A1: 2, B1: 0, C1: 0, E1: 10 })
        await set(s, 'D1', { formula: 'A1+100' })
        expect(await at(s, 'D1')).toEqual({ D1: 0 })
        // a value typed into the cycle breaks it: everything computes
        await set(s, 'B1', { value: 4 })
        expect(await at(s)).toEqual({ A1: 7, B1: 4, C1: 70, D1: 107, E1: 10 })
        expect(await held(s)).toEqual({})
        // closed again: the formula entered holds 0, the cells it brings into the cycle's way what they showed
        await set(s, 'B1', { formula: 'A1' })
        expect(await at(s)).toEqual({ A1: 7, B1: 0, C1: 70, D1: 107, E1: 10 })
        expect(await held(s)).toEqual({ A1: 7, C1: 70, D1: 107 })
        // undo and redo give the held values back
        await s.undo()
        expect(await at(s)).toEqual({ A1: 7, B1: 4, C1: 70, D1: 107, E1: 10 })
        await s.undo()
        expect(await at(s)).toEqual({ A1: 2, B1: 0, C1: 0, D1: 0, E1: 10 })
        expect(await held(s)).toEqual({ A1: 2 })
        await s.redo()
        expect(await at(s)).toEqual({ A1: 7, B1: 4, C1: 70, D1: 107, E1: 10 })
        await s.redo()
        expect(await at(s)).toEqual({ A1: 7, B1: 0, C1: 70, D1: 107, E1: 10 })
        expect(await held(s)).toEqual({ A1: 7, C1: 70, D1: 107 })
      } finally {
        await s.close()
      }
    })

    it('a formula typed closes a chain: the cells it brings into the way hold what they showed; nothing stored, they hold empty text', async () => {
      const s = await open(book({ A1: { value: 2 }, B1: { formula: 'A1' }, C1: { formula: 'B1*10' } }), { clock, execution })
      try {
        await s.setCellText('S', 'A1', '=B1/2+1')
        expect(await at(s, 'A1:C1')).toEqual({ A1: 0, B1: 2, C1: 20 })
      } finally {
        await s.close()
      }
      const t = await open(book(), { clock, execution, cached: { sheets: [] } })
      try {
        expect(await at(t, 'A1:C1')).toEqual({ A1: '', B1: '', C1: '' })
      } finally {
        await t.close()
      }
    })

    it('a look keeps the held value, a formula typed over it does not; previews change nothing', async () => {
      const s = await open(book(), { clock, execution, cached: STORED })
      try {
        await s.patchCells([{ sheet: 'S', cell: 'A1', set: { numberFormat: '0.00' } }])
        expect(await at(s)).toEqual({ A1: 2, B1: 2, C1: 0, E1: 6 })
        const hash = await s.getHash()
        expect((await s.previewChanges([{ op: 'setCells', edits: [{ sheet: 'S', cell: 'B1', input: { value: 4 } }] }])).changes.map((c) => [c.cell, c.value])).toEqual([['B1', 4], ['A1', 3], ['C1', 30]])
        expect((await s.previewCellText('S', 'B1', '9')).ok).toBe(true)
        expect([await s.getHash(), await held(s)]).toEqual([hash, { A1: 2, B1: 2 }])
        await s.setCellText('S', 'A1', '=B1/2+1')
        expect(await at(s)).toEqual({ A1: 0, B1: 2, C1: 0, E1: 6 })
        expect((await s.getBook()).sheets[0].cells.A1).toEqual({ formula: 'B1/2+1', numberFormat: '0.00' })
      } finally {
        await s.close()
      }
    })

    it('a row inserted moves the held values with their cells; a copy pasted into the way holds 0; the reason and the origin', async () => {
      const s = await open(book(), { clock, execution, cached: STORED })
      try {
        await s.editStructure({ op: 'insertRows', sheet: 'S', row: 1, count: 1 })
        expect(await at(s, 'A2:E2')).toEqual({ A2: 2, B2: 2, C2: 0, E2: 6 })
        // B2 = A2 holds 2; pasted, D2 = C2 reads the cycle too and holds 0
        await s.editStructure({ op: 'copyRange', from: { sheet: 'S', range: 'B2' }, to: { sheet: 'S', range: 'D2' } })
        expect(await at(s, 'D2')).toEqual({ D2: 0 })
        expect(await s.getCellIssue('S', 'C2')).toEqual({ error: 'CYCLE', reason: 'circular reference: not computed, the value is held', origin: { sheet: 'S', cell: 'A2' } })
        expect(await s.getCellIssue('S', 'E2')).toBeNull()
        await s.undo()
        await s.undo()
        expect(await at(s)).toEqual({ A1: 2, B1: 2, C1: 0, E1: 6 })
      } finally {
        await s.close()
      }
    })

    it('with iterative calculation the cycle is computed and held values are dropped', async () => {
      const s = await open({ ...book({ ...LOOP, A1: { formula: 'B1/2+1', held: 5 } }), iteration: {} }, { clock, execution, cached: STORED })
      try {
        expect((await at(s)).A1).toBeCloseTo(2, 2)
        expect(await held(s)).toEqual({})
      } finally {
        await s.close()
      }
    })
  })
})

describe.each(['the same thread', 'a thread'] as const)('a calculator computes with the stored values: %s', (kind) => {
  it('as compute does', async () => {
    const calc = kind === 'the same thread' ? createInlineCalc() : await keptPool()
    const r = await calc.compute(book(), { clock, cached: STORED })
    expect(r.hash).toBe(compute(book(), { clock, cached: STORED }).hash)
    expect(Object.fromEntries(Object.entries(r.sheets[0].cells).map(([k, c]) => [k, c.value]))).toEqual({ A1: 2, B1: 2, C1: 0, E1: 6, K1: 3 })
  })
})
