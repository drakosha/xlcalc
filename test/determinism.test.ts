/**
 * Детерминизм: хэш не зависит от порядка входа, числа прогонов, процесса и его часовой зоны
 * (formulajs работает с датами через JS Date в местной зоне); копейки не теряются.
 */
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { compute, type BookInput, type CellInput } from '../src/index.js'
import { mulberry32, synthBook } from '../src/synth.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const synth = synthBook({ seed: 1, rows: 340 })
const child = fileURLToPath(new URL('./fixtures/hash-child.ts', import.meta.url))

function shuffled(book: BookInput, seed: number): BookInput {
  const rnd = mulberry32(seed)
  const mix = <T>(xs: T[]) => {
    const a = [...xs]
    for (let i = a.length - 1; i > 0; i--) {
      const j = Math.floor(rnd() * (i + 1))
      ;[a[i], a[j]] = [a[j], a[i]]
    }
    return a
  }
  return { sheets: mix(book.sheets).map((s) => ({ name: s.name, cells: Object.fromEntries(mix(Object.entries(s.cells))) })), names: mix(book.names ?? []) }
}

function childHash(tz?: string): { hash: string; tz: string; offset: number } {
  const out = execFileSync(process.execPath, ['--import', 'tsx', child, '340'], { encoding: 'utf8', env: { ...process.env, ...(tz ? { TZ: tz } : {}) } })
  return JSON.parse(out.trim().split('\n').at(-1)!)
}

describe('детерминизм', () => {
  const base = compute(synth.book, { clock }).hash

  it('перемешанный вход и три прогона дают один хэш', () => {
    for (const seed of [1, 2, 3]) expect(compute(shuffled(synth.book, seed), { clock }).hash).toBe(base)
  })

  it('дочерний процесс даёт тот же хэш', () => {
    expect(childHash().hash).toBe(base)
  }, 60_000)

  it('часовая зона процесса не влияет: Лос-Анджелес, Москва, UTC', () => {
    const runs = ['America/Los_Angeles', 'Europe/Moscow', 'UTC'].map(childHash)
    expect(runs.map((r) => r.offset)).toEqual([480, -180, 0])
    for (const r of runs) expect(r.hash).toBe(base)
  }, 120_000)

  it('0.1+0.2 в выдаче 0.3; -0 выходит как 0', () => {
    const c = compute({ sheets: [{ name: 'S', cells: { A1: { formula: '0.1+0.2' }, A2: { formula: '-0*1' }, A3: { formula: 'A1=0.3' } } }] }, { clock }).sheets[0].cells
    expect(c.A1.value).toBe(0.3)
    expect(Object.is(c.A2.value, -0)).toBe(false)
    expect(c.A3.value).toBe(false) // внутри расчёт идёт с полной точностью, как в Excel
  })

  it('сумма 3 400 сумм с копейками, округлённая до копеек, равна сумме в копейках', () => {
    const rnd = mulberry32(5)
    const cells: Record<string, CellInput> = {}
    let kop = 0n
    for (let i = 1; i <= 3400; i++) {
      const k = Math.floor(rnd() * 2e11) - 5e10
      kop += BigInt(k)
      cells[`A${i}`] = { value: k / 100 }
    }
    cells.B1 = { formula: 'ROUND(SUM(A1:A3400),2)' }
    const v = compute({ sheets: [{ name: 'S', cells }] }, { clock }).sheets[0].cells.B1.value as number
    expect(BigInt(Math.round(v * 100))).toBe(kop)
  })
})
