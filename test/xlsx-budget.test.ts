/**
 * Speed of the xlsx subpath: a book of 20,000 rows × 40 values reads in 2 s at most (median of 3, in a fresh
 * process). The full book runs with `FORMULA_BENCH=1`, a book of 2,000 rows always (0.2 s).
 */
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const child = fileURLToPath(new URL('./fixtures/xlsx-bench-child.ts', import.meta.url))

function bench(rows: number) {
  const out = execFileSync(process.execPath, ['--expose-gc', '--import', 'tsx', child, String(rows)], { encoding: 'utf8' })
  return JSON.parse(out.trim().split('\n').at(-1)!) as { values: { readMs: number; writeMs: number; heapMB: number; sizeMB: number }; synth: { readMs: number; writeMs: number } }
}

describe('xlsx budgets', () => {
  it('a book of 2,000 rows × 40 values reads in 0.2 s at most', () => {
    expect(bench(2_000).values.readMs).toBeLessThanOrEqual(200)
  }, 120_000)

  it.skipIf(process.env.FORMULA_BENCH !== '1')('a book of 20,000 rows × 40 values reads in 2 s at most', () => {
    const r = bench(20_000)
    console.log(JSON.stringify(r))
    expect(r.values.readMs).toBeLessThanOrEqual(2_000)
  }, 300_000)
})
