import { describe, expect, it } from 'vitest'
import { compute, openSession, type BookInput, type CellInput } from '../src/index.js'
import { createState } from '../src/engine.js'
import { CalcError } from '../src/exec/budget.js'
import { Rng } from '../src/values.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

it('descending prefix sums reuse read work instead of reading a quadratic number of cells', () => {
  const n = 1024
  const state = createState({ sheets: [{ name: 'S', cells: Object.fromEntries(Array.from({ length: n }, (_, i) => [`A${i + 1}`, { value: 1 }])) }] }, { now: 0, seed: 1 })
  let reads = 0
  state.ctx.sheets[0].cols[0] = new Proxy(state.ctx.sheets[0].cols[0], {
    get(target, key, receiver) { if (typeof key === 'string' && /^\d+$/.test(key)) reads++; return Reflect.get(target, key, receiver) },
  })
  for (let end = n - 1; end >= 0; end--) expect(state.rt.F.SUM(new Rng(0, 0, 0, end, 0))).toBe(end + 1)
  expect(reads).toBeLessThan(n * 300)
})

describe.each(['closures', 'eval'] as const)('range aggregates, %s', (execution) => {
  it('keeps row-major addition order and does not replace an argument inside a multi-argument sum by its subtotal', () => {
    const cells: Record<string, CellInput> = { A1: { value: 1e16 }, A2: { value: 1 }, A3: { value: -1e16 } }
    for (let r = 4; r <= 400; r++) cells[`A${r}`] = { value: 0 }
    Object.assign(cells, { C1: { formula: 'SUM(A1:A400)' }, C2: { formula: 'SUM(1,A1:A400)' }, C3: { formula: 'SUM(A1:A2)' }, C4: { formula: 'SUM((A1:A2,A3:A400))' } })
    const v = compute({ sheets: [{ name: 'S', cells }] }, { clock, execution }).sheets[0].cells
    expect([v.C1.value, v.C2.value, v.C3.value, v.C4.value]).toEqual([0, 0, 1e16, 0])
  })

  it('retains an unrounded partial sum and distinguishes range errors from arithmetic overflow', () => {
    const cells: Record<string, CellInput> = {
      A1: { value: 1e-308 }, A256: { value: 1e-308 }, A257: { value: 1e-308 },
      B1: { value: 1e308 }, B2: { value: 1e308 }, B257: { formula: 'NA()' },
      D1: { formula: 'SUM(A1:A256)' }, D2: { formula: 'SUM(A1:A257)' },
      E1: { formula: 'SUM(B1:B256)' }, E2: { formula: 'SUM(B1:B257)' }, E3: { formula: 'SUM(B1:B257,1/0)' },
      F1: { formula: 'COUNT(B1:B257)' }, F2: { formula: 'MAX(B1:B257)' }, F3: { formula: 'MIN(B1:B257)' },
    }
    const v = compute({ sheets: [{ name: 'S', cells }] }, { clock, execution }).sheets[0].cells
    expect([v.D1.value, v.D2.value]).toEqual([0, 3e-308])
    expect([v.E1.value, v.E2.value, v.E3.value, v.F1.value, v.F2.value, v.F3.value]).toEqual([{ error: 'NUM' }, { error: 'NA' }, { error: 'DIV0' }, 2, { error: 'NA' }, { error: 'NA' }])
  })

  it('invalidates prefixes on edits, grows whole columns and keeps session/full hashes equal', () => {
    const cells: Record<string, CellInput> = {}
    for (let r = 1; r <= 1000; r++) { cells[`A${r}`] = { value: r }; cells[`B${r}`] = { formula: `SUM($A$1:A${r})` } }
    cells.D1 = { formula: 'SUM(A:A)' }; cells.D2 = { formula: 'COUNT(A:A)' }; cells.D3 = { formula: 'MIN(A:A)' }; cells.D4 = { formula: 'MAX(A:A)' }
    const s = openSession({ sheets: [{ name: 'S', cells }] }, { clock, execution, display: false })
    try {
      for (const [cell, input] of [['A500', { value: -10 }], ['A1', { value: 'text' }], ['A1500', { value: 42 }], ['A500', null]] as const) {
        s.setCells([{ sheet: 'S', cell, input }])
        expect(s.hash()).toBe(compute(s.book(), { clock, execution, display: false }).hash)
      }
      expect(s.restructure({ op: 'insertRows', sheet: 'S', row: 2, count: 2 }).ok).toBe(true)
      expect(s.hash()).toBe(compute(s.book(), { clock, execution, display: false }).hash)
    } finally { s.close() }
  })

  it('computes one range of more values than limits.cells, which a function that reads the range whole is refused', () => {
    const cells: Record<string, CellInput> = {}
    for (let r = 1; r <= 3000; r++) { cells[`A${r}`] = { value: r % 7 }; cells[`B${r}`] = { formula: `SUM($A$1:A${r})+COUNT($A$1:A${r})+MIN($A$1:A${r})+MAX($A$1:A${r})` } }
    const book: BookInput = { sheets: [{ name: 'S', cells }] }
    const limits = { cells: 1000 }
    expect(compute(book, { clock, execution, limits }).hash).toBe(compute(book, { clock, execution }).hash)
    const s = openSession(book, { clock, execution, display: false, limits })
    try {
      s.setCells([{ sheet: 'S', cell: 'A1500', input: { value: 100 } }])
      expect(s.hash()).toBe(compute(s.book(), { clock, execution, display: false }).hash)
    } finally { s.close() }
    for (const formula of ['AVERAGE(A1:A3000,1)', 'SUM(A1:A3000,1)', 'SUM(A1:A3000*1)']) {
      const e = (() => { try { compute({ sheets: [{ name: 'S', cells: { ...cells, D1: { formula } } }] }, { clock, execution, limits }) } catch (x) { return x } })()
      expect(e instanceof CalcError && e.limit, formula).toBe('cells')
    }
  })

  it('keeps horizontal prefixes and column extensions in their original addition order', () => {
    const state = createState({ sheets: [{ name: 'S', cells: {} }] }, { now: 0, seed: 1, execution })
    state.ctx.sheets[0].cols = Array.from({ length: 600 }, (_, c) => [c % 3 === 0 ? 1e16 : c % 3 === 1 ? 1 : -1e16, 0])
    state.ctx.sheets[0].rows = 2
    for (let end = 599; end >= 0; end--) {
      let expected = 0
      for (let c = 0; c <= end; c++) expected += state.ctx.sheets[0].cols[c][0] as number
      expect(state.rt.F.SUM(new Rng(0, 0, 0, 0, end))).toBe(expected)
    }
    const book: BookInput = { sheets: [{ name: 'S', cells: {
      A1: { value: 1e16 }, B1: { value: 1 }, A2: { value: -1e16 }, B2: { value: 1 },
      D1: { formula: 'SUM(A1:A256)' }, D2: { formula: 'SUM(A1:B256)' },
    } }] }
    expect(compute(book, { clock, execution }).sheets[0].cells.D2.value).toBe(1)
  })

  it('invalidates aggregate checkpoints when a spill grows, shrinks, is blocked and recovers', () => {
    const s = openSession({ sheets: [{ name: 'S', cells: {
      D1: { value: 512 }, A1: { formula: 'SEQUENCE(D1)' },
      B1: { formula: 'SUM(A1:A256)' }, B2: { formula: 'SUM(A1:A512)' },
      B3: { formula: 'COUNT(A:A)' }, B4: { formula: 'MIN(A:A)' }, B5: { formula: 'MAX(A:A)' },
    } }] }, { clock, execution, display: false })
    try {
      for (const [cell, input] of [['D1', { value: 600 }], ['D1', { value: 300 }], ['A257', { value: 'blocked' }], ['A257', null]] as const) {
        s.setCells([{ sheet: 'S', cell, input }])
        expect(s.hash()).toBe(compute(s.book(), { clock, execution, display: false }).hash)
      }
      expect(s.result().sheets[0].cells.B3.value).toBe(300)
    } finally { s.close() }
  })
})
