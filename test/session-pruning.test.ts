import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { compute, type AsyncSession, type BookInput, type CellInput } from '../src/index.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const chain = (): Record<string, CellInput> => {
  const cells: Record<string, CellInput> = { A1: { value: 1.1 }, A2: { formula: 'INT(A1)' } }
  for (let r = 3; r <= 96; r++) cells[`A${r}`] = { formula: `A${r - 1}+1` }
  return cells
}
afterAll(closeCalcs)

describe.each(['closures', 'eval'] as const)('pruned recalculation, %s', (execution) => {
  describe.each(KINDS)('%s', (kind) => {
    let open: Awaited<ReturnType<typeof opener>>
    beforeAll(async () => { open = await opener(kind) })
    const session = (cells: Record<string, CellInput>) => open({ sheets: [{ name: 'S', cells }] }, { clock, execution, display: false })
    const check = async (s: AsyncSession) => {
      const got = await s.getResult()
      const expected = compute(await s.getBook(), { clock: s.clock, execution, display: false })
      expect(got.hash).toBe(expected.hash)
      expect(got.sheets).toEqual(expected.sheets)
    }

    it('keeps independently dirty joins and propagates a later actual change', async () => {
      const s = await session({ ...chain(), Z1: { value: 0 }, C1: { formula: 'A96+Z1' } })
      try {
        const result = await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 1.9 } }, { sheet: 'S', cell: 'Z1', input: { value: 10 } }])
        expect(result.changes.map((c) => c.cell)).toEqual(['A1', 'Z1', 'C1'])
        expect((await s.getCell('S', 'C1'))!.value).toBe(105)
        await check(s)
        await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 2.1 } }])
        expect((await s.getCell('S', 'C1'))!.value).toBe(106)
        await check(s)
        await s.undo(); await check(s)
        await s.redo(); await check(s)
      } finally { await s.close() }
    })

    it('passes changed formula values through shared range helpers', async () => {
      const s = await session({ ...chain(), C1: { formula: 'SUM(A2:A96)' }, D1: { formula: 'C1*2' } })
      try {
        for (const value of [1.9, 2.1, 1.1]) {
          await s.setCells([{ sheet: 'S', cell: 'A1', input: { value } }])
          await check(s)
        }
      } finally { await s.close() }
    })

    it('propagates a new array with the same top-left value, including readers outside its anchor', async () => {
      const cells = chain()
      cells.A1 = { value: 0 }
      cells.D1 = { formula: 'IF(A1=0,1,SEQUENCE(A1,1,1,0))' }
      cells.E1 = { formula: 'SUM(D2:D3)' }
      cells.F1 = { formula: 'ROWS(D1#)' }
      const s = await session(cells)
      try {
        for (const value of [2, 1, 0, 3]) {
          await s.setCells([{ sheet: 'S', cell: 'A1', input: { value } }])
          await check(s)
        }
      } finally { await s.close() }
    })

    it('propagates a nested summary marker even when the replaced formula keeps its value', async () => {
      const cells: Record<string, CellInput> = { A1: { value: 3 }, A2: { value: 1 }, B1: { formula: 'A1+1' }, B2: { value: 1 }, C1: { formula: 'SUBTOTAL(9,B1:B2)' } }
      for (let r = 2; r <= 96; r++) cells[`C${r}`] = { formula: `C${r - 1}+1` }
      const s = await session(cells)
      try {
        await s.setCells([{ sheet: 'S', cell: 'B1', input: { formula: 'SUBTOTAL(9,A1:A2)' } }])
        expect((await s.getCell('S', 'C1'))!.value).toBe(1)
        await check(s)
        await s.undo(); await check(s)
        await s.redo(); await check(s)
      } finally { await s.close() }
    })
  })
})
