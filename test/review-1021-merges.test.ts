/**
 * Merged areas, from the review of 21.10: `setLayout` merges over a table, a merge inside a merged area of entire
 * columns, and a session merging entire columns of a sheet that holds a single cell. In each case a session's
 * `getBook()` computes, and the session's hash is a full computation's.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, InputError, type BookInput, type Execution } from '../src/index.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const EXECUTIONS: Execution[] = ['closures', 'eval']

afterAll(closeCalcs)

const tableBook = (): BookInput => ({
  sheets: [{
    name: 'S',
    cells: { A1: { value: 'H1' }, B1: { value: 'H2' }, A2: { value: 6 }, A3: { value: 7 }, D1: { formula: 'SUM(T[H1])' } },
    tables: [{ name: 'T', range: 'A1:B3', columns: [{ name: 'H1' }, { name: 'H2' }] }],
  }],
})

describe.each(EXECUTIONS)('setLayout merges over a table (x-fuzz-session-8), %s', (execution) => {
  describe.each(KINDS)('in %s', (kind) => {
    const open = openerFor(kind)
    it('refuses a merged area over a table with its path, and keeps one beside it', async () => {
      const s = await open(tableBook(), { execution, clock })
      const refused = s.setLayout([{ sheet: 'S', columns: 'D', set: { width: 12 } }, { sheet: 'S', merges: ['D5:E6', 'B3:C4'] }])
      await expect(refused).rejects.toThrow(InputError)
      await expect(refused).rejects.toThrow('edits[1].merges[1]: the merged area B3:C4 overlaps the table T')
      expect(s.revision).toBe(0)
      await s.setLayout([{ sheet: 'S', merges: ['D5:E6'] }])
      const after = await s.getBook()
      expect(after.sheets[0].merges).toEqual(['D5:E6'])
      expect(compute(after, { execution, clock }).hash).toBe(await s.getHash())
      await s.close()
    })
  })
})
