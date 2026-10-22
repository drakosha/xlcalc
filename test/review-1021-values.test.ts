/**
 * A copy from a filtered sheet pastes values, from the review of 21.10: the cells of a large spill past its stored part
 * paste their value too (0 for `A:A*2` below the data), as the same copy of a stored spill does.
 */
import { describe, expect, it } from 'vitest'
import { createSession, editStructure, type BookInput, type StructureOp } from '../src/index.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

/** S filtered with row 2 hidden, B1 spilling `formula` over the data A1:A3. */
const filtered = (formula: string): BookInput => ({
  sheets: [
    { name: 'S', filtered: true, rows: { '2': { hidden: true } }, cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, B1: { formula } } },
    { name: 'T', cells: {} },
  ],
})
const copy: StructureOp = { op: 'copyRange', from: { sheet: 'S', range: 'B1:B8' }, to: { sheet: 'T', range: 'A1' } }
const WANT = { A1: { value: 2 }, A2: { value: 6 }, A3: { value: 0 }, A4: { value: 0 }, A5: { value: 0 }, A6: { value: 0 }, A7: { value: 0 } }

describe('a copy from a filtered sheet over a large spill (fix-u295)', () => {
  it('pastes the values of its unstored cells, as of a stored spill', () => {
    for (const formula of ['A1:A8*2', 'A:A*2']) {
      const r = editStructure(filtered(formula), copy)
      if (!r.ok) throw new Error(r.message)
      expect(r.book.sheets[1].cells, formula).toEqual(WANT)
    }
  })

  it.each(['closures', 'eval'] as const)('in a session (%s)', (execution) => {
    const s = createSession(filtered('A:A*2'), { execution, clock })
    expect(s.editStructure(copy).ok).toBe(true)
    expect(s.getBook().sheets[1].cells).toEqual(WANT)
  })
})
