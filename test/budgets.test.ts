/**
 * Engine budgets on the synthetic book: cold compute ≤ 1.5 s (median), memory ≤ 300 MB (heapUsed; rss
 * is only recorded), class A edit ≤ 100 ms (p95), validating 3,000 formulas ≤ 300 ms. The full book
 * (3,400 rows) runs with `FORMULA_BENCH=1`, the short one (340 rows) always.
 */
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const child = fileURLToPath(new URL('./fixtures/bench-child.ts', import.meta.url))

function bench(rows: number) {
  const out = execFileSync(process.execPath, ['--expose-gc', '--import', 'tsx', child, String(rows)], { encoding: 'utf8' })
  return JSON.parse(out.trim().split('\n').at(-1)!) as {
    formulas: number
    coldMedianMs: number
    heapUsedMB: number
    rssMB: number
    editA: { p95: number }
    editB: { median: number }
    validateMedianMs: number
  }
}

function check(r: ReturnType<typeof bench>) {
  expect(r.coldMedianMs).toBeLessThanOrEqual(1500)
  expect(r.heapUsedMB).toBeLessThanOrEqual(300)
  expect(r.editA.p95).toBeLessThanOrEqual(100)
  // the budget is 300 ms per 3,000 formulas; the synthetic book has 3,907
  expect(r.validateMedianMs).toBeLessThanOrEqual((300 * r.formulas) / 3000)
}

describe('budgets', () => {
  it('short book (340 rows)', () => check(bench(340)), 120_000)
  it.skipIf(process.env.FORMULA_BENCH !== '1')('full book (3,400 rows)', () => {
    const r = bench(3400)
    console.log(JSON.stringify(r))
    check(r)
  }, 300_000)
})
