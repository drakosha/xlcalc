/** LINEST, TREND, LOGEST and GROWTH on data far from 0 against their spread, against exact least squares on the stored values. */
import { describe, expect, it } from 'vitest'
import { compute, type BookInput, type CellValue } from '../src/index.js'

const clock = { now: '2026-10-11T12:00:00Z', seed: 7 }
const close = (actual: CellValue, expected: number) => {
  expect(typeof actual).toBe('number')
  expect(Math.abs((actual as number) - expected)).toBeLessThanOrEqual(1e-12 * Math.abs(expected))
}

/** Columns from row 1 (A, B, …) and formulas in column Z. */
const bookOf = (cols: number[][], formulas: string[]): BookInput => ({
  sheets: [{ name: 'S', cells: Object.fromEntries([
    ...cols.flatMap((c, j) => c.map((value, i) => [`${String.fromCharCode(65 + j)}${i + 1}`, { value }])),
    ...formulas.map((formula, i) => [`Z${i + 1}`, { formula }]),
  ]) }],
})

const K = Array.from({ length: 12 }, (_, k) => k)
/** Residuals of the lines below: they sum to 0, and Σk·e = −5.25. */
const E = [0.5, -0.25, 0, 0.75, -0.5, 0.25, 0, -0.75, 0.5, -0.25, 0.25, -0.5]

describe.each(['closures', 'eval'] as const)('least squares far from 0, %s', (execution) => {
  const values = (cols: number[][], formulas: string[]) => {
    const cells = compute(bookOf(cols, formulas), { clock, execution, display: false }).sheets[0].cells
    return formulas.map((_, i) => cells[`Z${i + 1}`].value)
  }

  it('keeps the slope of 10,000 values 1E12 + 0, 0.01, 0.02 on their steps, which lie on a line', () => {
    // stored as 1E12 + 82/8192 · step exactly
    const steps = Array.from({ length: 10000 }, (_, i) => i % 3)
    const v = values([steps.map((s) => 1e12 + s / 100), steps], ['INDEX(LINEST(A1:A10000,B1:B10000),1,1)', 'INDEX(LINEST(A1:A10000,B1:B10000),1,2)-1E12', 'TREND(A1:A10000,B1:B10000,3)-1E12', 'INDEX(LINEST(A1:A10000,B1:B10000,TRUE,TRUE),3,1)'])
    close(v[0], 82 / 8192)
    expect(v[1]).toBe(0)
    expect(v[2]).toBe(246 / 8192)
    expect(v[3]).toBe(1)
  })

  it('keeps a line on x = 1E10 + 0…11, its fitted value past the data and the standard error of its intercept', () => {
    const v = values([K.map((k) => 1e10 + k), K.map((k) => 3 + 2 * k + E[k])], [
      'INDEX(LINEST(B1:B12,A1:A12),1,1)', 'TREND(B1:B12,A1:A12,1E10+12)', 'INDEX(TREND(B1:B12,A1:A12),1)', 'INDEX(LINEST(B1:B12,A1:A12,TRUE,TRUE),2,2)',
    ])
    close(v[0], 2 - 5.25 / 143)
    close(v[1], 27 - 34.125 / 143)
    close(v[2], 3 + 28.875 / 143)
    // the residual sum of squares 2.18225524475524… over 10 degrees of freedom, (XᵀX)⁻¹ by Fractions
    close(v[3], 390647234.77693009)
  })

  it('keeps a growth on x = 1E10 + 0…11', () => {
    const grow = K.map((k) => Math.round(100 * 1.1 ** k * 100) / 100)
    const v = values([K.map((k) => 1e10 + k), grow], ['INDEX(LOGEST(B1:B12,A1:A12),1,1)', 'GROWTH(B1:B12,A1:A12,1E10+12)'])
    // least squares on ln y by Decimal at 60 digits
    close(v[0], 1.0999987487542413)
    close(v[1], 313.83972734359537)
  })

  it('keeps two variables near 1E6 and y near 1E9', () => {
    const y = K.map((k) => 1e9 + k / 2 - ((k * k) % 7) / 4 + E[k] / 8)
    const v = values([K.map((k) => 1e6 + k), K.map((k) => 1e6 + ((k * k) % 7)), y], [
      'INDEX(LINEST(C1:C12,A1:B12),1,1)', 'INDEX(LINEST(C1:C12,A1:B12),1,2)', 'INDEX(LINEST(C1:C12,A1:B12),1,3)-1E9', 'INDEX(LINEST(C1:C12,A1:B12,TRUE,TRUE),3,2)',
    ])
    close(v[0], -0.2468650865879661)
    close(v[1], 0.49522449815383013)
    close(v[2], -248359.39130918795)
    close(v[3], 0.061352792084726382)
  })
})
