/** API conveniences: the default clock of `compute` and `openSession`. */
import { describe, expect, it } from 'vitest'
import { compute, openSession, type BookInput } from '../src/index.js'

const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: 'NOW()' }, A2: { formula: 'RAND()' }, A3: { value: 1 }, A4: { formula: 'A3*2' } } }] }

describe('default clock', () => {
  it('compute without a clock uses the current moment and a random seed; its clock reproduces the result', () => {
    const r = compute(book)
    expect(Math.abs(Date.parse(r.clock.now) - Date.now())).toBeLessThan(60_000)
    expect(Number.isInteger(r.clock.seed) && r.clock.seed >= 0 && r.clock.seed < 2 ** 32).toBe(true)
    const again = compute(book, { clock: r.clock })
    expect(again.hash).toBe(r.hash)
    expect(again.sheets[0].cells.A2.value).toBe(r.sheets[0].cells.A2.value)
    expect(compute(book).clock.seed).not.toBe(r.clock.seed)
  })

  it('openSession without options keeps the clock of opening; a full computation with it gives the same hash', () => {
    const s = openSession(book)
    const clock = s.clock
    s.setCells([{ sheet: 'S', cell: 'A3', input: { value: 5 } }])
    expect(s.clock).toEqual(clock)
    const edited: BookInput = { sheets: [{ name: 'S', cells: { ...book.sheets[0].cells, A3: { value: 5 } } }] }
    expect(s.result().hash).toBe(compute(edited, { clock }).hash)
    expect(s.setClock({ now: '2026-01-01T00:00:00Z', seed: 1 }).length).toBeGreaterThan(0)
  })
})
