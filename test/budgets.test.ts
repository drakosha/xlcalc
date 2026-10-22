/**
 * Engine budgets on the synthetic book, in both execution modes: cold compute ≤ 1.5 s (median), memory
 * ≤ 300 MB (heapUsed; rss is only recorded), class A edit ≤ 100 ms (p95), validating 3,000 formulas
 * ≤ 300 ms. The full book (3,400 rows) runs with `FORMULA_BENCH=1`. The short one (340 rows) runs always, against
 * limits of its own, so that it catches a regression of a few times rather than of an order of magnitude: on 22.10.2026
 * it measured cold 110 to 140 ms, heap 40 MB, edit p95 1.5 ms and validation 40 to 50 ms alone, and with 16 copies
 * running at once on 18 cores up to 280 ms, 23 ms and 125 ms; the limits are about twice the loaded figures.
 */
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import type { Execution } from '../src/contract.js'

const child = fileURLToPath(new URL('./fixtures/bench-child.ts', import.meta.url))

function bench(rows: number, execution: Execution) {
  const out = execFileSync(process.execPath, ['--expose-gc', '--import', 'tsx', child, String(rows), execution], { encoding: 'utf8' })
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

const FULL = { coldMs: 1500, heapMB: 300, editP95Ms: 100, validateMsPer3000: 300 }
const SHORT = { coldMs: 500, heapMB: 100, editP95Ms: 40, validateMsPer3000: 200 }

function check(r: ReturnType<typeof bench>, limits: typeof FULL) {
  expect(r.coldMedianMs).toBeLessThanOrEqual(limits.coldMs)
  expect(r.heapUsedMB).toBeLessThanOrEqual(limits.heapMB)
  expect(r.editA.p95).toBeLessThanOrEqual(limits.editP95Ms)
  // the validation budget is per 3,000 formulas; the short book has 3,907
  expect(r.validateMedianMs).toBeLessThanOrEqual((limits.validateMsPer3000 * r.formulas) / 3000)
}

describe.each(['closures', 'eval'] as const)('budgets, execution %s', (execution) => {
  it('short book (340 rows)', () => check(bench(340, execution), SHORT), 120_000)
  it.skipIf(process.env.FORMULA_BENCH !== '1')('full book (3,400 rows)', () => {
    const r = bench(3400, execution)
    console.log(JSON.stringify(r))
    check(r, FULL)
  }, 300_000)
})
