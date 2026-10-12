import { describe, expect, it } from 'vitest'
import { applyEdits, createState } from '../src/engine.js'
import { run } from '../src/exec/budget.js'
import type { CellInput } from '../src/contract.js'

describe.each(['closures', 'eval'] as const)('incremental work, %s', (execution) => {
  it('does not reread a shared unchanged short range for every recomputed formula', () => {
    const n = 256, cells: Record<string, CellInput> = {}
    for (let r = 1; r <= 20; r++) cells[`D${r}`] = { value: r }
    for (let r = 1; r <= n; r++) {
      cells[`A${r}`] = { value: r }
      cells[`B${r}`] = { formula: `A${r}+SUM($D$1:$D$20)` }
    }
    const state = createState({ sheets: [{ name: 'S', cells }] }, { now: 0, seed: 1, execution })
    let reads = 0
    state.ctx.sheets[0].cols[3] = new Proxy(state.ctx.sheets[0].cols[3], {
      get(target, k, receiver) {
        if (typeof k === 'string' && /^\d+$/.test(k)) reads++
        return Reflect.get(target, k, receiver)
      },
    })
    run(applyEdits(state, Array.from({ length: n }, (_, i) => ({ sheet: 'S', cell: `A${i + 1}`, input: { value: i + 2 } }))))
    expect(state.rt.cell(0, n - 1, 1)).toBe(n + 211)
    expect(reads).toBeLessThan(n * 4)
    run(applyEdits(state, [{ sheet: 'S', cell: 'D1', input: { value: 11 } }]))
    expect(state.rt.cell(0, n - 1, 1)).toBe(n + 221)
  })

  it('bounds the replay of cached prefixes after a middle input changes', () => {
    const n = 1024, cells: Record<string, CellInput> = {}
    for (let r = 1; r <= n; r++) {
      cells[`A${r}`] = { value: 1 }
      cells[`B${r}`] = { formula: `AVERAGE($A$1:A${r})+COUNTIF($A$1:A${r},">0")` }
    }
    const state = createState({ sheets: [{ name: 'S', cells }] }, { now: 0, seed: 1, execution })
    let reads = 0
    state.ctx.sheets[0].cols[0] = new Proxy(state.ctx.sheets[0].cols[0], {
      get(target, k, receiver) {
        if (typeof k === 'string' && /^\d+$/.test(k)) reads++
        return Reflect.get(target, k, receiver)
      },
    })
    run(applyEdits(state, [{ sheet: 'S', cell: 'A512', input: { value: 2 } }]))
    expect(state.rt.cell(0, n - 1, 1)).toBe(n + 1 + 1 / n)
    expect(reads).toBeLessThan(n * 96)
  })

  it.each([false, true])('finds the readers of a contiguous input batch without quadratic range checks, reverse=%s', (reverse) => {
    const n = 512, cells: Record<string, CellInput> = {}
    for (let r = 1; r <= n; r++) {
      cells[`A${r}`] = { value: r }
      cells[`B${r}`] = { formula: `A${r}*3` }
    }
    const state = createState({ sheets: [{ name: 'S', cells }] }, { now: 0, seed: 1, execution })
    let probes = 0
    const scan = state.listeners.forEach.bind(state.listeners)
    state.listeners.forEach = (area, visit) => scan(new Proxy(area, {
      get(target, key, receiver) {
        if (key === 'r1' || key === 'r2') probes++
        return Reflect.get(target, key, receiver)
      },
    }), visit)
    const edits = Array.from({ length: n }, (_, i) => ({ sheet: 'S', cell: `A${i + 1}`, input: { value: i + 2 } }))
    if (reverse) edits.reverse()
    run(applyEdits(state, edits))
    expect(state.rt.cell(0, n - 1, 1)).toBe((n + 1) * 3)
    expect(probes).toBeLessThan(n * 32)
  })
})
