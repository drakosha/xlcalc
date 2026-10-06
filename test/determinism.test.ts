/**
 * Determinism: the hash does not depend on input order, the number of runs, the process or its time zone
 * (dates are computed in UTC, never through the local zone); cents are not lost.
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

describe('determinism', () => {
  const base = compute(synth.book, { clock }).hash

  it('shuffled input and three runs give one hash', () => {
    for (const seed of [1, 2, 3]) expect(compute(shuffled(synth.book, seed), { clock }).hash).toBe(base)
  })

  it('a child process gives the same hash', () => {
    expect(childHash().hash).toBe(base)
  }, 60_000)

  it('the process time zone does not matter: Los Angeles, Moscow, UTC', () => {
    const runs = ['America/Los_Angeles', 'Europe/Moscow', 'UTC'].map(childHash)
    expect(runs.map((r) => r.offset)).toEqual([480, -180, 0])
    for (const r of runs) expect(r.hash).toBe(base)
  }, 120_000)

  it('0.1+0.2 comes out as 0.3; -0 comes out as 0', () => {
    const c = compute({ sheets: [{ name: 'S', cells: { A1: { formula: '0.1+0.2' }, A2: { formula: '-0*1' }, A3: { formula: 'A1=0.3' } } }] }, { clock }).sheets[0].cells
    expect(c.A1.value).toBe(0.3)
    expect(Object.is(c.A2.value, -0)).toBe(false)
    expect(c.A3.value).toBe(false) // internally full precision is kept, as in Excel
  })

  it('the sum of 3,400 amounts with cents, rounded to cents, equals the sum in cents', () => {
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
