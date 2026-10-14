import { afterAll, describe, expect, it } from 'vitest'
import { applyEdits, createState } from '../src/engine.js'
import { run } from '../src/exec/budget.js'
import { compute, type AsyncSession, type BookInput, type CellEdit } from '../src/index.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
afterAll(closeCalcs)

describe.each(['closures', 'eval'] as const)('formula storage, %s', (execution) => {
  it('bounds storage across repeated replacement of one formula', () => {
    const state = createState({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, B1: { formula: 'A1+1' } } }] }, { now: 0, seed: 1, execution })
    for (let i = 0; i < 256; i++) run(applyEdits(state, [{ sheet: 'S', cell: 'B1', input: { formula: `A1+${i % 2 + 1}` } }]))
    expect(state.rt.cell(0, 0, 1)).toBe(3)
    expect(state.formulas.length).toBeLessThanOrEqual(2)
    expect(state.succ.length).toBeLessThanOrEqual(2)
  })

  it('reuses holes and releases empty trailing storage', () => {
    const state = createState({ sheets: [{ name: 'S', cells: { A1: { formula: '1' }, A2: { formula: '2' }, A3: { formula: '3' } } }] }, { now: 0, seed: 1, execution })
    run(applyEdits(state, [{ sheet: 'S', cell: 'A2', input: { value: 2 } }, { sheet: 'S', cell: 'B1', input: { formula: 'SUM(A1:A3)' } }]))
    expect(state.rt.cell(0, 0, 1)).toBe(6)
    expect(state.formulas.length).toBeLessThanOrEqual(3)
    run(applyEdits(state, ['A1', 'A3', 'B1'].map((cell) => ({ sheet: 'S', cell, input: null }))))
    expect(state.formulas.length).toBe(0)
    expect(state.succ.length).toBe(0)
  })

  describe.each(KINDS)('%s', (kind) => {
    const open = openerFor(kind)
    it('keeps references, spills, cycles, nested summaries, clock and history after replacements', async () => {
      const book: BookInput = { sheets: [{ name: 'S', cells: {
        A1: { value: 3 }, A2: { value: 1 }, B1: { formula: 'A1+1' }, B2: { value: 1 },
        C1: { formula: 'B1*2' }, D1: { formula: 'SEQUENCE(A1)' }, E1: { formula: 'SUM(D1:D5)' }, F1: { formula: 'SUBTOTAL(9,B1:B2)' },
      } }] }
      const s = await open(book, { clock, execution, display: false })
      const check = async (s: AsyncSession) => {
        const result = await s.getResult()
        const full = compute(await s.getBook(), { clock: s.clock, execution, display: false })
        expect(result.hash).toBe(full.hash)
        expect(result.sheets).toEqual(full.sheets)
      }
      const steps: CellEdit[][] = [
        [{ sheet: 'S', cell: 'B1', input: { formula: 'SUBTOTAL(9,A1:A2)' } }],
        [{ sheet: 'S', cell: 'B1', input: { formula: 'A1+1' } }],
        [{ sheet: 'S', cell: 'D1', input: { formula: 'SEQUENCE(A1)+10' } }],
        [{ sheet: 'S', cell: 'D1', input: null }],
        [{ sheet: 'S', cell: 'D1', input: { formula: 'SEQUENCE(A1)' } }],
        [{ sheet: 'S', cell: 'B1', input: { formula: 'C1+1' } }],
        [{ sheet: 'S', cell: 'B1', input: { formula: 'A1+1' } }],
        [{ sheet: 'S', cell: 'B1', input: { formula: 'NOW()' } }],
      ]
      try {
        for (const edits of steps) { await s.setCells(edits); await check(s) }
        await s.setClock({ ...clock, now: '2027-01-01T00:00:00Z' })
        await check(s)
        for (let i = 0; i < 4; i++) { await s.undo(); await check(s) }
        for (let i = 0; i < 4; i++) { await s.redo(); await check(s) }
      } finally { await s.close() }
    })
  })
})
