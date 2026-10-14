import { describe, expect, it } from 'vitest'
import { applyEdits, createState } from '../src/engine.js'
import { run } from '../src/exec/budget.js'
import type { CellInput } from '../src/contract.js'
import { columnReads, evaluations } from './engine-counts.js'

describe.each(['closures', 'eval'] as const)('incremental work, %s', (execution) => {
  it('prunes an unchanged branch when an independent root changes', () => {
    const n = 1024, cells: Record<string, CellInput> = { A1: { value: 1.1 }, A2: { formula: 'INT(A1)' } }
    for (let r = 3; r <= n; r++) cells[`A${r}`] = { formula: `A${r - 1}+1` }
    cells.C1 = { value: 1 }; cells.C2 = { formula: 'C1+1' }
    const state = createState({ sheets: [{ name: 'S', cells }] }, { now: 0, seed: 1, execution })
    const evaluated = evaluations(state)
    run(applyEdits(state, [{ sheet: 'S', cell: 'A1', input: { value: 1.9 } }, { sheet: 'S', cell: 'C1', input: { value: 2 } }]))
    expect(state.rt.cell(0, n - 1, 0)).toBe(n - 1)
    expect(state.rt.cell(0, 1, 2)).toBe(3)
    expect(evaluated.n).toBe(2)
  })

  it('keeps the graph when a scalar formula keeps its dependencies', () => {
    const state = createState({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, B1: { formula: 'A1+1' }, C1: { formula: 'B1*2' } } }] }, { now: 0, seed: 1, execution })
    const { succ, listeners, topo } = state
    run(applyEdits(state, [{ sheet: 'S', cell: 'B1', input: { formula: 'A1+2' } }]))
    expect(state.rt.cell(0, 0, 2)).toBe(6)
    expect(state.succ).toBe(succ)
    expect(state.listeners).toBe(listeners)
    expect(state.topo).toBe(topo)
  })

  it('does not reread lookup keys when only the result column changes', () => {
    const cells: Record<string, CellInput> = {}
    for (let r = 1; r <= 512; r++) {
      cells[`A${r}`] = { value: r }; cells[`B${r}`] = { value: r * 3 }
    }
    for (let r = 1; r <= 16; r++) cells[`D${r}`] = { formula: `VLOOKUP(${512 - r},$A$1:$B$512,2,FALSE)` }
    const state = createState({ sheets: [{ name: 'S', cells }] }, { now: 0, seed: 1, execution })
    const reads = columnReads(state, 0)
    run(applyEdits(state, [{ sheet: 'S', cell: 'B511', input: { value: 9 } }]))
    expect(state.rt.cell(0, 0, 3)).toBe(9)
    expect(reads()).toBeLessThan(32)
  })

  it('stops a scalar chain when the first formula keeps its value', () => {
    const n = 1024, cells: Record<string, CellInput> = { A1: { value: 1.1 }, A2: { formula: 'INT(A1)' } }
    for (let r = 3; r <= n; r++) cells[`A${r}`] = { formula: `A${r - 1}+1` }
    const state = createState({ sheets: [{ name: 'S', cells }] }, { now: 0, seed: 1, execution })
    const evaluated = evaluations(state)
    run(applyEdits(state, [{ sheet: 'S', cell: 'A1', input: { value: 1.9 } }]))
    expect(state.rt.cell(0, n - 1, 0)).toBe(n - 1)
    expect(evaluated.n).toBeLessThan(4)
    evaluated.n = 0
    run(applyEdits(state, [{ sheet: 'S', cell: 'A1', input: { value: 2.1 } }]))
    expect(state.rt.cell(0, n - 1, 0)).toBe(n)
    expect(evaluated.n).toBe(n - 1)
  })

  it('releases spill round bookkeeping after a calculation', () => {
    const cells: Record<string, CellInput> = { A1: { value: 1 } }
    for (let r = 2; r <= 1024; r++) cells[`A${r}`] = { formula: `A${r - 1}+1` }
    const state = createState({ sheets: [{ name: 'S', cells }] }, { now: 0, seed: 1, execution })
    expect(state.spills.done.size).toBe(0)
    run(applyEdits(state, [{ sheet: 'S', cell: 'A1', input: { value: 2 } }]))
    expect(state.rt.cell(0, 1023, 0)).toBe(1025)
    expect(state.spills.done.size).toBe(0)
  })

  it('does not reread a shared unchanged short range for every recomputed formula', () => {
    const n = 256, cells: Record<string, CellInput> = {}
    for (let r = 1; r <= 20; r++) cells[`D${r}`] = { value: r }
    for (let r = 1; r <= n; r++) {
      cells[`A${r}`] = { value: r }
      cells[`B${r}`] = { formula: `A${r}+SUM($D$1:$D$20)` }
    }
    const state = createState({ sheets: [{ name: 'S', cells }] }, { now: 0, seed: 1, execution })
    const reads = columnReads(state, 3)
    run(applyEdits(state, Array.from({ length: n }, (_, i) => ({ sheet: 'S', cell: `A${i + 1}`, input: { value: i + 2 } }))))
    expect(state.rt.cell(0, n - 1, 1)).toBe(n + 211)
    expect(reads()).toBeLessThan(n * 4)
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
    const reads = columnReads(state, 0)
    run(applyEdits(state, [{ sheet: 'S', cell: 'A512', input: { value: 2 } }]))
    expect(state.rt.cell(0, n - 1, 1)).toBe(n + 1 + 1 / n)
    expect(reads()).toBeLessThan(n * 96)
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
