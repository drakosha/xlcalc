/**
 * `registerPivot` in a pool already running: the pool was made, and its threads started with sessions of a workbook
 * without PivotTables, before the engine was loaded; the threads hear it (`onPivotLoaded`, `worker/pool.ts`), so that
 * `addPivotTable` and a refresh lay the PivotTable out in a session opened before and in one opened after.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { registerPivot, type BookInput, type Calc } from '../src/index.js'
import { pivotEngine } from '../src/pivot/index.js'
import { pool } from './fixtures/pool.js'

const clock = { now: '2026-10-31T12:00:00Z', seed: 1 }

const book = (): BookInput => ({
  sheets: [
    { name: 'Data', cells: { A1: { value: 'K' }, B1: { value: 'N' }, A2: { value: 'a' }, B2: { value: 2 }, A3: { value: 'b' }, B3: { value: 3 }, A4: { value: 'a' }, B4: { value: 5 } } },
    { name: 'P', cells: {} },
  ],
})

const laid = [
  ['Row Labels', 'Sum of N'],
  ['a', 7],
  ['b', 3],
  ['Grand Total', 10],
]

const calcs: Calc[] = []
afterAll(() => Promise.all(calcs.map((c) => c.close())))

describe('registerPivot in a pool already running', () => {
  it('lays PivotTables out in its threads started before, in sessions opened before and after', async () => {
    const calc = await pool({ threads: 2 })
    calcs.push(calc)
    // both threads busy with a workbook without PivotTables before the engine is loaded
    const before = [await calc.createSession(book(), { clock }), await calc.createSession(book(), { clock })]
    expect(await before[0].editStructure({ op: 'refreshPivots' })).toMatchObject({ ok: false, code: 'pivot', message: expect.stringContaining('registerPivot') })
    registerPivot(pivotEngine)
    for (const s of before) {
      const r = await s.editStructure({ op: 'addPivotTable', sheet: 'P', cell: 'A3', source: { type: 'range', sheet: 'Data', range: 'A1:B4' }, definition: { rows: [0], values: [{ field: 1 }] } })
      expect(r).toMatchObject({ ok: true })
      expect(r.ok && r.stale).toBeUndefined()
      expect(await s.getValues('P', 'A3:B6')).toEqual(laid)
      await s.setCells([{ sheet: 'Data', cell: 'B3', input: { value: 4 } }])
      expect(await s.editStructure({ op: 'refreshPivots' })).toMatchObject({ ok: true })
      expect(await s.getValues('P', 'B5:B6')).toEqual([[4], [11]])
    }
    // a session opened after, of the workbook with the PivotTable, refreshes it in the same threads
    const after = await calc.createSession(await before[0].getBook(), { clock })
    await after.setCells([{ sheet: 'Data', cell: 'B2', input: { value: 1 } }])
    expect(await after.editStructure({ op: 'refreshPivots' })).toMatchObject({ ok: true })
    expect(await after.getValues('P', 'A3:B6')).toEqual([['Row Labels', 'Sum of N'], ['a', 6], ['b', 4], ['Grand Total', 10]])
    await Promise.all([...before, after].map((s) => s.close()))
  })
})
