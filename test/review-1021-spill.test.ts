/** Reproductions from the review of 21.10 on spills and legacy arrays: books that settle, areas a session keeps whole. */
import { describe, expect, it } from 'vitest'
import { compute, createInlineCalc, createSession, type BookInput, type CellInput } from '../src/index.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const book = (cells: Record<string, CellInput>, extra: Partial<BookInput> = {}): BookInput => ({ sheets: [{ name: 'S', cells }], ...extra })

/** The book computed, opened as a session and entered cell by cell in both orders: each session's hash is a full computation's. */
async function sameEverywhere(b: BookInput, execution: 'closures' | 'eval') {
  const full = compute(b, { clock, execution })
  expect(createSession(b, { clock, execution }).getHash()).toBe(full.hash)
  const calc = createInlineCalc({ sliceMs: 1 })
  const inline = await calc.createSession(b, { clock, execution })
  expect(await inline.getHash()).toBe(full.hash)
  await inline.close()
  const cells = Object.entries(b.sheets[0].cells)
  for (const order of [cells, [...cells].reverse()]) {
    const s = createSession({ ...b, sheets: [{ ...b.sheets[0], cells: {} }] }, { clock, execution })
    for (const [cell, input] of order) s.setCells([{ sheet: 'S', cell, input }])
    expect(s.getHash()).toBe(compute(s.getBook(), { clock, execution }).hash)
  }
  return full.sheets[0].cells
}

describe.each(['closures', 'eval'] as const)('execution %s', (execution) => {
  it('an array that reads a blocked array and spills over its range settles (x-fuzz-modes-4)', async () => {
    const c = await sameEverywhere(book({ A1: { formula: 'SEQUENCE(2)' }, A2: { formula: 'A1+{1;2}' } }), execution)
    expect([c.A1.value, c.A2.value, c.A3.value]).toEqual([{ error: 'SPILL' }, { error: 'SPILL' }, { error: 'SPILL' }])
    const d = await sameEverywhere(book({ B12: { formula: 'SEQUENCE(3,3)' }, D13: { formula: 'SUM(B12)+{1;2;3}' } }), execution)
    expect([d.B12.value, d.D13.value, d.D15.value]).toEqual([{ error: 'SPILL' }, { error: 'SPILL' }, { error: 'SPILL' }])
    const e = await sameEverywhere(book({ A1: { formula: 'SEQUENCE(2)' }, A2: { formula: 'IFERROR(A1,0)+{1;2}' } }), execution)
    expect([e.A1.value, e.A2.value, e.A3.value]).toEqual([{ error: 'SPILL' }, 1, 2])
  })

  it('a reader of an array in a cycle through its spill settles under iterative calculation (x-fuzz-modes-4)', async () => {
    const cells = { G4: { formula: 'SUM(E7:F10)' }, E8: { formula: 'IF(D6:E9>365,D6:F8,7)' } }
    // each iteration starts from 0: in one, G4 comes before E8 spills, in many it reads the spill
    expect((await sameEverywhere(book(cells, { iteration: { maxIterations: 1 } }), execution)).G4.value).toBe(0)
    expect((await sameEverywhere(book(cells, { iteration: {} }), execution)).G4.value).toBe(42)
  })
})
