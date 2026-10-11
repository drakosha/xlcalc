/**
 * A structure operation's geometry through a session, done, undone and redone: `moved` and `mapCell` of an undo and a
 * redo are those of the operation taken back and again, where the operation did more than its arguments say (visible
 * rows of a filtered sheet, a table deleted with its cells). In a synchronous session, in the same thread and in a
 * thread, with closures and eval.
 */
import { afterAll, describe, expect, it } from 'vitest'
import type { AsyncSession, BookInput, CellRef, Execution } from '../src/index.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

/** Values in A1:A9 of a filtered sheet whose rows 3, 5 and 6 are hidden, a sum over them, and a table on T. */
const book = (): BookInput => ({
  sheets: [
    {
      name: 'S',
      cells: { ...Object.fromEntries([...Array(9).keys()].map((i) => [`A${i + 1}`, { value: i + 1 }])), C1: { formula: 'SUM(A1:A9)' } },
      filtered: true,
      rows: { '3': { hidden: true }, '5:6': { hidden: true } },
    },
    {
      name: 'T',
      cells: { A1: { value: 'Qty' }, A2: { value: 3 }, A3: { value: 4 }, C1: { formula: 'SUM(Tab[Qty])' } },
      tables: [{ name: 'Tab', range: 'A1:A3', columns: [{ name: 'Qty' }] }],
    },
  ],
})

const where = (mapCell: ((ref: CellRef) => CellRef | null) | undefined, cells: string[], sheet = 'S') => cells.map((cell) => mapCell!({ sheet, cell })?.cell ?? null)

describe.each(['closures', 'eval'] as const)('the geometry of an operation done, undone and redone, %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    const open = async (): Promise<AsyncSession> => (await opener(kind))(book(), { clock, execution })

    it('rows deleted on a filtered sheet: the visible ones, run by run; undone and redone the same way', async () => {
      const s = await open()
      // rows 2 to 7 hold the visible rows 2, 4 and 7
      const r = await s.restructure({ op: 'deleteRows', sheet: 'S', row: 2, count: 6 })
      if (!r.ok) throw new Error(r.message)
      const runs = [7, 4, 2].map((at) => ({ kind: 'lines', sheet: 'S', axis: 'rows', at, count: -1 }))
      expect(r.moved).toEqual(runs)
      expect(where(r.mapCell, ['A1', 'A2', 'A3', 'A7', 'A8'])).toEqual(['A1', null, 'A2', null, 'A5'])
      const u = (await s.undo())!
      expect(u.moved).toEqual([2, 4, 7].map((at) => ({ kind: 'lines', sheet: 'S', axis: 'rows', at, count: 1 })))
      expect(where(u.mapCell, ['A1', 'A2', 'A4', 'A5', 'A6'])).toEqual(['A1', 'A3', 'A6', 'A8', 'A9'])
      const again = (await s.redo())!
      expect(again.moved).toEqual(runs)
      expect(where(again.mapCell, ['A1', 'A2', 'A3', 'A7', 'A8'])).toEqual(where(r.mapCell, ['A1', 'A2', 'A3', 'A7', 'A8']))
      await s.close()
    })

    it("a table deleted with its cells: a redo's mapCell takes them away as the operation's did", async () => {
      const s = await open()
      const r = await s.restructure({ op: 'deleteTable', table: 'Tab' })
      if (!r.ok) throw new Error(r.message)
      expect(where(r.mapCell, ['A1', 'A3', 'C1'], 'T')).toEqual([null, null, 'C1'])
      // undone, the table's cells come from nowhere, as those of a paste undone
      const u = (await s.undo())!
      expect(where(u.mapCell, ['A1', 'C1'], 'T')).toEqual([null, 'C1'])
      const again = (await s.redo())!
      expect(where(again.mapCell, ['A1', 'A3', 'C1'], 'T')).toEqual([null, null, 'C1'])
      // the cells the redo took away need no change to be gone
      expect(again.changes.filter((c) => c.sheet === 'T' && c.cell !== 'C1')).toEqual([])
      await s.close()
    })
  })
})
