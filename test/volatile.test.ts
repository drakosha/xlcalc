/** Volatile functions are computed from the clock: the moment for TODAY/NOW, the seed and the address for RAND. */
import { describe, expect, it } from 'vitest'
import { compute, type BookInput } from '../src/index.js'

const book: BookInput = {
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { formula: 'TODAY()' },
        A2: { formula: 'NOW()' },
        A3: { formula: 'RAND()' },
        A4: { formula: 'RAND()' },
        A5: { formula: 'RANDBETWEEN(1,6)' },
        B1: { formula: 'A1+1' },
        B2: { formula: 'B1*2' },
        C1: { formula: '"RAND()"&"TODAY()"' },
        C2: { value: 5 },
      },
    },
  ],
}
const clock = { now: '2025-01-01T18:00:00Z', seed: 42 }

describe('volatile functions', () => {
  it('TODAY and NOW from the moment (UTC)', () => {
    const c = compute(book, { clock }).sheets[0].cells
    expect(c.A1.value).toBe(45658)
    expect(c.A2.value).toBe(45658.75)
    expect(compute(book, { clock: { now: '2025-01-01T23:30:00-03:00', seed: 42 } }).sheets[0].cells.A1.value).toBe(45659)
  })

  it('the same clock gives one hash over any number of runs', () => {
    const hashes = new Set([1, 2, 3, 4, 5].map(() => compute(book, { clock }).hash))
    expect(hashes.size).toBe(1)
  })

  it('RAND from the seed and the cell address', () => {
    const a = compute(book, { clock }).sheets[0].cells
    const b = compute(book, { clock: { ...clock, seed: 43 } }).sheets[0].cells
    expect(a.A3.value).not.toBe(a.A4.value)
    expect(a.A3.value).not.toBe(b.A3.value)
    for (const v of [a.A3.value, a.A4.value]) expect(v as number).toBeGreaterThanOrEqual(0)
    expect([1, 2, 3, 4, 5, 6]).toContain(a.A5.value)
  })

  it('cells depending on volatile functions are listed; a call inside a string literal does not count', () => {
    const r = compute(book, { clock })
    expect(r.volatile.map((v) => v.cell)).toEqual(['A1', 'B1', 'A2', 'B2', 'A3', 'A4', 'A5'])
    expect(r.clock).toEqual(clock)
  })
})
