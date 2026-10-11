import { describe, expect, it } from 'vitest'
import { compute, createSession, type CellInput } from '../src/index.js'
import { createState } from '../src/engine.js'
import { Rng } from '../src/values.js'
import { CacheBudget } from '../src/range-cache.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

it.each(['AVERAGE', 'COUNTA', 'COUNTIF'])('%s reuses descending prefixes without rereading their whole input', (name) => {
  const n = 1024, cells = Object.fromEntries(Array.from({ length: n }, (_, i) => [`A${i + 1}`, { value: 1 }]))
  const state = createState({ sheets: [{ name: 'S', cells }] }, { now: 0, seed: 1 })
  let reads = 0
  state.ctx.sheets[0].cols[0] = new Proxy(state.ctx.sheets[0].cols[0], {
    get(target, key, receiver) { if (typeof key === 'string' && /^\d+$/.test(key)) reads++; return Reflect.get(target, key, receiver) },
  })
  for (let end = n - 1; end >= 0; end--) {
    const g = new Rng(0, 0, 0, end, 0)
    expect(name === 'COUNTIF' ? state.rt.F.COUNTIF(g, '>0') : state.rt.F[name](g)).toBe(name === 'AVERAGE' ? 1 : end + 1)
  }
  expect(reads).toBeLessThan(n * 300)
})

it('COUNTIF retains shared equality indexes for many criteria over the same range', () => {
  const n = 1024, cells = Object.fromEntries(Array.from({ length: n }, (_, i) => [`A${i + 1}`, { value: i }]))
  const state = createState({ sheets: [{ name: 'S', cells }] }, { now: 0, seed: 1 })
  let reads = 0
  state.ctx.sheets[0].cols[0] = new Proxy(state.ctx.sheets[0].cols[0], {
    get(target, key, receiver) { if (typeof key === 'string' && /^\d+$/.test(key)) reads++; return Reflect.get(target, key, receiver) },
  })
  for (let value = 0; value < 100; value++) expect(state.rt.F.COUNTIF(new Rng(0, 0, 0, n - 1, 0), value)).toBe(1)
  expect(reads).toBeLessThan(n * 10)
})

it('does not materialize an optional equality index larger than the cache budget', () => {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= 1024; r++) cells[`A${r}`] = { value: r }
  for (let r = 1; r <= 100; r++) cells[`D${r}`] = { formula: `COUNTIF($A$1:$A$1024,${r})` }
  const state = createState({ sheets: [{ name: 'S', cells }] }, { now: 0, seed: 1, cache: new CacheBudget(8192) })
  expect(state.ctx.index.size).toBe(0)
  expect(state.ctx.flat.size).toBe(0)
  for (let r = 0; r < 100; r++) expect(state.rt.cell(0, r, 3)).toBe(1)
})

describe.each(['closures', 'eval'] as const)('prefix statistics, %s', (execution) => {
  it('streams single ranges beyond the allocation limit and still bounds materializing fallbacks', () => {
    const cells: Record<string, CellInput> = {}
    for (let r = 1; r <= 3000; r++) cells[`A${r}`] = { value: r % 7 }
    for (const [i, formula] of ['AVERAGE(A1:A3000)', 'COUNTA(A1:A3000)', 'COUNTIF(A1:A3000,">0")', 'COUNTIFS(A1:A3000,">0")', 'COUNTIF(A1:A3000,1)', 'COUNTIF(A1:A3000,2)', 'COUNTIF(A1:A3000,3)'].entries()) cells[`C${i + 1}`] = { formula }
    const book = { sheets: [{ name: 'S', cells }] }, opts = { clock, execution, display: false }
    expect(compute(book, { ...opts, limits: { cells: 1000 } }).hash).toBe(compute(book, opts).hash)
    expect(() => compute({ sheets: [{ name: 'S', cells: { ...cells, C1: { formula: 'AVERAGE(A1:A3000,1)' } } }] }, { ...opts, limits: { cells: 1000 } })).toThrow(/over the limit/)
  })

  it('preserves compensated means, constant extremes, overflow fallback and reference error order', () => {
    const cells: Record<string, CellInput> = { A1: { value: 1e16 }, A2: { value: 1 }, A3: { value: -1e16 }, B1: { value: 1e308 }, B2: { value: 1e308 }, B257: { value: -1e308 } }
    for (let r = 4; r <= 300; r++) cells[`A${r}`] = { value: r === 257 ? 'text' : 0 }
    const formulas = ['AVERAGE(A1:A300)', 'AVERAGE(A1:A256)', 'COUNTA(A1:A300)', 'AVERAGE(B1:B256)', 'AVERAGE(B1:B257)']
    for (const [i, formula] of formulas.entries()) cells[`D${i + 1}`] = { formula }
    const s = createSession({ sheets: [{ name: 'S', cells }] }, { clock, execution, display: false })
    try {
      const values = s.getCells('S', 'D1:D5')
      expect(values.D1.value).toBe(Number((1 / 299).toPrecision(15)))
      expect(values.D2.value).toBe(1 / 256)
      expect(values.D3.value).toBe(300)
      expect(values.D4.value).toBe(1e308)
      expect(values.D5.value).toBe(Number((1e308 / 3).toPrecision(15)))
      for (const [cell, input] of [['A128', { formula: 'NA()' }], ['A129', { formula: '1/0' }], ['A128', null], ['A257', { value: false }]] as const) {
        s.setCells([{ sheet: 'S', cell, input }])
        expect(s.getHash()).toBe(compute(s.getBook(), { clock, execution, display: false }).hash)
      }
    } finally { s.close() }
  })

  it('COUNTIF keeps indexed equality, blank tails and year-dependent text matching', () => {
    const cells: Record<string, CellInput> = {
      A1: { value: '1/31' }, A2: { value: ' 1/31' }, A3: { value: '' }, A4: { value: false }, A5: { formula: 'NA()' },
      C1: { formula: 'COUNTIF(A1:A300,DATE(2026,1,31))' }, C2: { formula: 'COUNTIF(A1:A300," 1/31")' },
      C3: { formula: 'COUNTIF(A:A,"")' }, C4: { formula: 'COUNTIF(A:A,NA())' }, C5: { formula: 'COUNTIF(A1:A300,FALSE)' },
    }
    const s = createSession({ sheets: [{ name: 'S', cells }] }, { clock, execution, display: false })
    try {
      expect(Object.values(s.getCells('S', 'C1:C5')).map((c) => c.value)).toEqual([1, 1, 1048572, 1, 1])
      const next = { ...clock, now: '2027-10-04T12:00:00Z' }
      s.setClock(next)
      expect(s.getCells('S', 'C1').C1.value).toBe(0)
      expect(s.getHash()).toBe(compute(s.getBook(), { clock: next, execution, display: false }).hash)
    } finally { s.close() }
  })
})
