import { describe, expect, it } from 'vitest'
import { applyEdits, createState } from '../src/engine.js'
import { run } from '../src/exec/budget.js'
import type { CellInput } from '../src/contract.js'

describe.each(['closures', 'eval'] as const)('incremental work, %s', (execution) => {
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
