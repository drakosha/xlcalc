import { describe, expect, it } from 'vitest'
import { compute, createSession, type BookInput, type CellValue } from '../src/index.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const close = (actual: CellValue, expected: number) => {
  expect(typeof actual).toBe('number')
  if (expected === 0) expect(actual).toBe(0)
  else expect(Math.abs((actual as number) / expected - 1)).toBeLessThan(2e-13)
}
const bookOf = (xs: number[], ys: number[], formulas: string[]): BookInput => ({ sheets: [{ name: 'S', cells: {
  ...Object.fromEntries(xs.map((value, i) => [`A${i + 1}`, { value }])),
  ...Object.fromEntries(ys.map((value, i) => [`B${i + 1}`, { value }])),
  ...Object.fromEntries(formulas.map((formula, i) => [`D${i + 1}`, { formula }])),
} }] })

describe.each(['closures', 'eval'] as const)('stable statistics, %s', (execution) => {
  it('keeps a small spread around a large mean, against an exact rational reference', () => {
    const xs = Array.from({ length: 10000 }, (_, i) => 1e12 + (i % 3) / 100)
    const book = bookOf(xs, [], ['AVERAGE(A:A)', 'VAR.P(A:A)', 'STDEV.P(A:A)', 'DEVSQ(A:A)'])
    const v = compute(book, { clock, execution, display: false }).sheets[0].cells
    // Exact binary values are 1e12 + {0, 82/8192, 164/8192}, with counts {3334,3333,3333}.
    const n = 10000n, sum = 3333n * (82n + 164n), squares = 3333n * (82n ** 2n + 164n ** 2n)
    const variance = Number(n * squares - sum * sum) / Number(n * n * 8192n ** 2n)
    expect(v.D1.value).toBe(1e12 + 0.01)
    close(v.D2.value, variance); close(v.D3.value, Math.sqrt(variance)); close(v.D4.value, variance * 10000)
    const s = createSession(book, { clock, execution })
    s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 1e12 + 0.02 } }])
    expect(s.getHash()).toBe(compute(s.getBook(), { clock, execution }).hash)
    s.close()
  })

  it.each([1e150, 1e-200])('does not lose a representable standard deviation at scale %s', (scale) => {
    const v = compute(bookOf([scale, scale * 2], [1, 2], ['STDEV.P(A:A)', 'STDEV.S(A:A)', 'CORREL(A:A,B:B)', 'SLOPE(A:A,B:B)']), { clock, execution }).sheets[0].cells
    close(v.D1.value, scale / 2); close(v.D2.value, scale / Math.sqrt(2)); close(v.D3.value, 1); close(v.D4.value, scale)
  })

  it('preserves Excel\'s numeric error when the variance itself overflows', () => {
    const v = compute(bookOf([1e300, -1e300], [], ['STDEV.P(A:A)', 'STDEV.S(A:A)']), { clock, execution }).sheets[0].cells
    expect([v.D1.value, v.D2.value]).toEqual([{ error: 'NUM' }, { error: 'NUM' }])
  })

  it('a constant near the largest double has zero variance and a finite mean', () => {
    const v = compute(bookOf([1e308, 1e308, 1e308], [], ['AVERAGE(A:A)', 'VAR.P(A:A)', 'STDEV.P(A:A)', 'DEVSQ(A:A)']), { clock, execution }).sheets[0].cells
    close(v.D1.value, 1e308)
    for (const a of ['D2', 'D3', 'D4']) close(v[a].value, 0)
  })

  it('covariance keeps contributions across widely different scales', () => {
    const v = compute(bookOf([1e308, -1e308, 1e-200], [0, 0, 1e308], ['COVARIANCE.P(A:A,B:B)', 'COVARIANCE.S(A:A,B:B)']), { clock, execution }).sheets[0].cells
    close(v.D1.value, (2 / 9) * 1e108); close(v.D2.value, 1e108 / 3)
  })

  it('STEYX retains a small nonzero residual near a straight line', () => {
    const y = [1, 2.00000001, 3]
    const v = compute(bookOf([1, 2, 3], y, ['STEYX(B:B,A:A)', 'SLOPE(B:B,A:A)']), { clock, execution }).sheets[0].cells
    close(v.D1.value, (y[1] - 2) * Math.sqrt(2 / 3)); close(v.D2.value, 1)
  })

  it('retains a small intercept after cancellation of the two mean terms', () => {
    const y = [0.5, 1.00000001, 1.5]
    const v = compute(bookOf([0.25, 0.5, 0.75], y, ['INTERCEPT(B:B,A:A)', 'FORECAST(0,B:B,A:A)']), { clock, execution }).sheets[0].cells
    close(v.D1.value, (y[1] - 1) / 3); close(v.D2.value, (y[1] - 1) / 3)
  })

  it('STEYX near a strong correlation agrees with an independent Fraction/Decimal calculation', () => {
    const x = [1.218261304197592e-10, -1.25634210083295e-11, 1.4598980988767263e-10]
    const y = [6.150589335415155e-11, 1.9303787081095429e-10, 3.8973379524949484e-11]
    close(compute(bookOf(x, y, ['STEYX(B:B,A:A)']), { clock, execution }).sheets[0].cells.D1.value, 7.176370251649892e-13)
  })

  it('database variance and deviation use the same stable calculation', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: {
      A1: { value: 'x' }, A2: { value: 1e-200 }, A3: { value: 2e-200 }, C1: { value: 'x' },
      E1: { formula: 'DSTDEVP(A1:A3,"x",C1:C2)' }, E2: { formula: 'DVARP(A1:A3,"x",C1:C2)' },
    } }] }
    const v = compute(book, { clock, execution }).sheets[0].cells
    close(v.E1.value, 5e-201); expect(v.E2.value).toBe(0)
  })
})
