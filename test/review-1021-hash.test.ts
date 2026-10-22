/**
 * `inputHash` takes the order of the sheets when it is not their order by name, since a 3-D reference covers the sheets
 * between its ends in book order (fix-u385); a book whose sheets are in name order hashes as before.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, type BookInput, type CellInput } from '../src/index.js'
import { closeCalcs, KINDS, openerFor } from './session-kinds.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
/** Sheets A, B and C in the given order; A!B1 sums A1 over the sheets from A to B. */
const book = (order: string[]): BookInput => ({
  sheets: order.map((name) => ({ name, cells: (name === 'A' ? { A1: { value: 1 }, B1: { formula: 'SUM(A:B!A1)' } } : { A1: { value: name === 'B' ? 10 : 100 } }) as Record<string, CellInput> })),
})
// the inputHash of the sheets A, B, C before the order counted
const SORTED = '15f9c8393fe9030c60232d343635cbc714aa6104b34401f3a012a19a6c8d95e6'

describe.each(['closures', 'eval'] as const)('the order of the sheets in inputHash (fix-u385), %s', (execution) => {
  afterAll(closeCalcs)

  it('books that differ only in sheet order and compute differently get different inputHash', () => {
    const run = (order: string[]) => compute(book(order), { clock, execution })
    const [abc, acb, cab] = [run(['A', 'B', 'C']), run(['A', 'C', 'B']), run(['C', 'A', 'B'])]
    expect([abc, acb, cab].map((r) => r.sheets.find((s) => s.name === 'A')!.cells.B1.value)).toEqual([11, 111, 11])
    expect(new Set([abc.inputHash, acb.inputHash, cab.inputHash]).size).toBe(3)
    expect(abc.inputHash).toBe(SORTED)
  })

  describe.each(KINDS)('in %s', (kind) => {
    const open = openerFor(kind)
    it('moveSheet gives the inputHash of the moved book, and moving back the first one', async () => {
      const s = await open(book(['A', 'B', 'C']), { clock, execution })
      await s.editStructure({ op: 'moveSheet', sheet: 'C', before: 'B' })
      const moved = await s.getResult()
      expect(moved.inputHash).toBe(compute(book(['A', 'C', 'B']), { clock, execution }).inputHash)
      expect(moved.inputHash).not.toBe(SORTED)
      await s.editStructure({ op: 'moveSheet', sheet: 'B', before: 'C' })
      expect((await s.getResult()).inputHash).toBe(SORTED)
      await s.close()
    })
  })
})
