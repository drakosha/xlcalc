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

  it('RANDARRAY: each element from the seed, the cell and its position; RAND keeps its value', () => {
    const cells = { A1: { formula: 'RAND()' }, B1: { formula: 'INDEX(RANDARRAY(3),1)' }, B2: { formula: 'INDEX(RANDARRAY(3),2)' }, B3: { formula: 'SUM(MOD(RANDARRAY(4,5,1,6,TRUE),1))' }, B4: { formula: 'MIN(RANDARRAY(4,5,1,6,TRUE))>=1' }, B5: { formula: 'RANDARRAY(2)' } }
    const run = (seed: number) => compute({ sheets: [{ name: 'S', cells }] }, { clock: { ...clock, seed } })
    const a = run(42)
    const v = (r: typeof a, cell: string) => r.sheets[0].cells[cell].value
    // the value RAND had before RANDARRAY drew from the same function
    expect(v(a, 'A1')).toBe(0.0295275952666998)
    expect(v(a, 'B1')).not.toBe(v(a, 'B2'))
    expect([v(a, 'B3'), v(a, 'B4')]).toEqual([0, true])
    // RANDARRAY(2) spills into B6
    for (const cell of ['B5', 'B6']) expect(v(a, cell)).toBeTypeOf('number')
    expect(run(42).hash).toBe(a.hash)
    expect(v(run(43), 'B1')).not.toBe(v(a, 'B1'))
    expect(a.volatile.map((c) => c.cell)).toEqual(['A1', 'B1', 'B2', 'B3', 'B4', 'B5'])
  })

  it('cells depending on volatile functions are listed; a call inside a string literal does not count', () => {
    const r = compute(book, { clock })
    expect(r.volatile.map((v) => v.cell)).toEqual(['A1', 'B1', 'A2', 'B2', 'A3', 'A4', 'A5'])
    expect(r.clock).toEqual(clock)
  })
})
