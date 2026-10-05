/**
 * Бюджеты движка на синтетике (`docs/design-workbook.md`): холодный расчёт ≤ 1,5 с (медиана),
 * память ≤ 300 МБ (по heapUsed, rss только фиксируется), правка класса А ≤ 100 мс (p95), проверка
 * 3 000 формул ≤ 300 мс. Полная книга (3 400 строк) при `FORMULA_BENCH=1`, короткая (340) всегда.
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
  // бюджет 300 мс на 3 000 формул, у синтетики 3 907
  expect(r.validateMedianMs).toBeLessThanOrEqual((300 * r.formulas) / 3000)
}

describe('бюджеты', () => {
  it('короткая книга (340 строк)', () => check(bench(340)), 120_000)
  it.skipIf(process.env.FORMULA_BENCH !== '1')('полная книга (3 400 строк)', () => {
    const r = bench(3400)
    console.log(JSON.stringify(r))
    check(r)
  }, 300_000)
})
