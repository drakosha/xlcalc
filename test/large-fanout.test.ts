/**
 * Lists longer than a call's argument limit: one volatile cell read by 139,999 formulas, `FREQUENCY` over 130,000
 * numbers. Spreading such a list into one call overflows the stack; the engine must compute them as any other book.
 */
import { describe, expect, it } from 'vitest'
import { compute, createSession, type BookInput } from '../src/index.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const later = { now: '2026-10-05T12:00:00Z', seed: 7 }
const ROWS = 140_000

const fanout = (): BookInput => {
  const cells: BookInput['sheets'][number]['cells'] = { A1: { formula: 'NOW()' } }
  for (let i = 2; i <= ROWS; i++) cells[`A${i}`] = { formula: 'A1' }
  return { sheets: [{ name: 'S', cells }] }
}

describe.each(['closures', 'eval'] as const)('lists past the argument limit, %s', (execution) => {
  it('a volatile cell read by 139,999 formulas: compute and a session after setClock', () => {
    const book = fanout()
    const r = compute(book, { clock, execution, display: false })
    const v = r.sheets[0].cells
    expect(v.A1.value).toBe(46299.5)
    expect(v[`A${ROWS}`].value).toBe(46299.5)
    expect(r.volatile).toHaveLength(ROWS)
    const s = createSession(book, { clock, execution, display: false })
    expect(s.setClock(later).changes).toHaveLength(ROWS)
    expect(s.getResult().sheets[0].cells[`A${ROWS}`].value).toBe(46300.5)
    expect(s.getHash()).toBe(compute(book, { clock: later, execution, display: false }).hash)
    s.close()
  }, 120_000)

  it('FREQUENCY over 130,000 numbers', () => {
    const n = 130_000
    const cells: BookInput['sheets'][number]['cells'] = { C1: { formula: `FREQUENCY(A1:A${n},{4})` }, D1: { formula: 'SUM(FREQUENCY(A:A,{4}))' } }
    for (let i = 1; i <= n; i++) cells[`A${i}`] = { value: i % 10 }
    const v = compute({ sheets: [{ name: 'S', cells }] }, { execution, display: false }).sheets[0].cells
    expect([v.C1.value, v.C2.value, v.D1.value]).toEqual([65_000, 65_000, n])
  }, 120_000)
})
