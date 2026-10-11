import { describe, expect, it } from 'vitest'
import { compute, createSession, type BookInput, type CellInput } from '../src/index.js'
import { applyEdits, createState } from '../src/engine.js'
import { run } from '../src/exec/budget.js'
import { formatAddress } from '../src/model.js'
import { mulberry32 } from '../src/synth.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }

it('drops affected aggregate prefixes once before a large formula batch', () => {
  const n = 1024, cells: Record<string, CellInput> = { D1: { value: 1 } }
  for (let r = 1; r <= n; r++) { cells[`A${r}`] = { formula: '$D$1' }; cells[`B${r}`] = { formula: `SUM($A$1:A${r})` } }
  const state = createState({ sheets: [{ name: 'S', cells }] }, { now: 0, seed: 1 })
  let probes = 0
  const invalidate = state.ctx.aggregates.invalidate.bind(state.ctx.aggregates)
  state.ctx.aggregates.invalidate = (area) => invalidate(new Proxy(area, {
    get(target, key, receiver) { if (key === 'r1' || key === 'r2') probes++; return Reflect.get(target, key, receiver) },
  }))
  run(applyEdits(state, [{ sheet: 'S', cell: 'D1', input: { value: 2 } }]))
  expect(state.rt.cell(0, n - 1, 1)).toBe(n * 2)
  expect(probes).toBeLessThan(n * 100)
})

it.each([false, true])('compresses overlapping formula ranges, horizontal=%s', (horizontal) => {
  const n = 1024, cells: Record<string, CellInput> = {}
  for (let i = 0; i < n; i++) {
    cells[formatAddress(horizontal ? 0 : i, horizontal ? i : 0)] = { formula: '1' }
    cells[formatAddress(horizontal ? 1 : i, horizontal ? i : 1)] = { formula: `SUM(A1:${formatAddress(horizontal ? 0 : i, horizontal ? i : 0)})` }
  }
  const state = createState({ sheets: [{ name: 'S', cells }] }, { now: 0, seed: 1 })
  expect(state.succ.reduce((n, next) => n + next.length, 0)).toBeLessThan(n * 32)
  expect(state.rt.cell(0, horizontal ? 1 : n - 1, horizontal ? n - 1 : 1)).toBe(n)
})

describe.each(['closures', 'eval'] as const)('compressed graph, %s', (execution) => {
  it('has the same formula reachability as a direct reference graph over rectangles', () => {
    const rnd = mulberry32(73), cells: Record<string, CellInput> = {}
    for (let r = 0; r < 32; r++) for (let c = 0; c < 16; c++) cells[formatAddress(r, c)] = { formula: '1' }
    for (let i = 0; i < 80; i++) {
      const r = Math.floor(rnd() * 20), c = Math.floor(rnd() * 10)
      cells[formatAddress(i, 20)] = { formula: `SUM(${formatAddress(r, c)}:${formatAddress(r + 8, c + 5)})` }
    }
    const state = createState({ sheets: [{ name: 'S', cells }] }, { now: 0, seed: 1, execution })
    for (const from of state.formulas.filter((f) => f && f.c < 16)) {
      const reached = new Set<number>(), stack = [...state.succ[from!.id]]
      while (stack.length) {
        const id = stack.pop()!
        if (reached.has(id)) continue
        reached.add(id)
        for (const next of state.succ[id]) stack.push(next)
      }
      const actual = [...reached].filter((id) => state.formulas[id]).sort((a, b) => a - b)
      const expected = state.formulas.filter((f) => f?.refs.some((g) => g.s === from!.s && g.r1 <= from!.r && from!.r <= g.r2 && g.c1 <= from!.c && from!.c <= g.c2)).map((f) => f!.id)
      expect(actual).toEqual(expected)
    }
  })

  it('detects and breaks a cycle inside a shared range without marking an unrelated prefix cyclic', () => {
    const cells: Record<string, CellInput> = {}
    for (let r = 1; r <= 64; r++) cells[`A${r}`] = { formula: '1' }
    cells.A32 = { formula: 'B1' }; cells.B1 = { formula: 'SUM(A1:A64)' }; cells.B2 = { formula: 'SUM(A1:A16)' }
    const opts = { clock, execution, display: false }, s = createSession({ sheets: [{ name: 'S', cells }] }, opts)
    try {
      expect(s.getCells('S', 'B1:B2')).toMatchObject({ B1: { value: { error: 'CYCLE' } }, B2: { value: 16 } })
      s.setCells([{ sheet: 'S', cell: 'A32', input: { value: 2 } }])
      expect(s.getCells('S', 'B1').B1.value).toBe(65)
      expect(s.getHash()).toBe(compute(s.getBook(), opts).hash)
    } finally { s.close() }
  })

  it('propagates volatility and edits through range vertices', () => {
    const cells: Record<string, CellInput> = {}
    for (let r = 1; r <= 64; r++) cells[`A${r}`] = { formula: 'RAND()' }
    cells.B1 = { formula: 'SUM(A1:A64)' }; cells.B2 = { formula: 'B1*2' }
    const book: BookInput = { sheets: [{ name: 'S', cells }] }, s = createSession(book, { clock, execution, display: false })
    try {
      const next = { ...clock, seed: 11 }
      s.setClock(next)
      expect(s.getHash()).toBe(compute(s.getBook(), { clock: next, execution, display: false }).hash)
    } finally { s.close() }
  })
})
