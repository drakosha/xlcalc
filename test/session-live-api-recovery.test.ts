import { describe, expect, it } from 'vitest'
import { inside } from './fixtures/pool.js'
const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
describe.each(['closures', 'eval'] as const)('live query recovery, %s', (execution) => {
  it('keeps dependency pagination stable when a worker restores formula slots', async () => {
    const p = await inside({ threads: 1, snapshotEvery: 1 })
    const s = await p.calc.createSession({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, B1: { formula: 'A1+1' }, C1: { formula: 'A1+2' } } }] }, { clock, execution })
    try {
      await s.clearCells([{ sheet: 'S', range: 'B1' }], 'all')
      await s.setCells([{ sheet: 'S', cell: 'B1', input: { formula: 'A1+1' } }])
      const refs = [{ sheet: 'S', range: 'A1' }]
      const first = await s.getDependents(refs, { maxCells: 1 })
      p.crash(0)
      const second = await s.getDependents(refs, { maxCells: 1, cursor: first.cursor })
      const cells = [...first.dependents, ...second.dependents].map((c) => c.cell)
      expect(new Set(cells).size).toBe(2)
      expect(cells).toEqual(['C1', 'B1'])
    } finally {
      await s.close()
      await p.calc.close()
    }
  })
})
