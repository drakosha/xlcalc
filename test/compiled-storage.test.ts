import { describe, expect, it } from 'vitest'
import { applyEdits, createState, debugSource } from '../src/engine.js'
import { run } from '../src/exec/budget.js'

describe.each(['closures', 'eval'] as const)('compiled storage, %s', (execution) => {
  it('eval shares executable bodies, closures keep one each; each formula keeps its coordinates and literals', () => {
    const cells = Object.fromEntries(Array.from({ length: 512 }, (_, i) => [`A${i + 1}`, { value: i + 1 }])) as Record<string, { value?: number; formula?: string }>
    for (let r = 1; r <= 512; r++) cells[`B${r}`] = { formula: `A${r}+${r}` }
    const state = createState({ sheets: [{ name: 'S', cells }] }, { now: 0, seed: 1, execution })
    expect(state.rt.cell(0, 511, 1)).toBe(1024)
    const { bodies } = state.compiler as typeof state.compiler & { bodies: number }
    if (execution === 'eval') expect(bodies).toBeLessThan(8)
    else expect(bodies).toBe(512)
    run(applyEdits(state, [{ sheet: 'S', cell: 'A1', input: { value: 9 } }]))
    expect(state.rt.cell(0, 0, 1)).toBe(10)
    expect(state.rt.cell(0, 511, 1)).toBe(1024)
  })

  it('retains executable functions without a second copy of their generated text', () => {
    const state = createState({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { value: 2 }, B1: { formula: 'SUM(A1:A2)' }, C1: { formula: 'LET(x,B1,x*2)' } } }] }, { now: 0, seed: 1, execution })
    expect(state.formulas.every((f) => f?.source === undefined)).toBe(true)
    const source = debugSource(state)
    expect(source).toContain("// 'S'!B1 = SUM(A1:A2)")
    expect(source).toContain('R(0,0,0,1,0)')
    expect(source).toContain('F.SUM(')
    expect(source).toContain('v0=')
    run(applyEdits(state, [{ sheet: 'S', cell: 'A1', input: { value: 3 } }]))
    expect(state.rt.cell(0, 0, 2)).toBe(10)
    expect(state.formulas.every((f) => f?.source === undefined)).toBe(true)
    run(applyEdits(state, [{ sheet: 'S', cell: 'B1', input: { formula: 'A1*2' } }]))
    expect(state.rt.cell(0, 0, 2)).toBe(12)
    expect(debugSource(state)).toContain('mul(cell(0,0,0),2)')
    expect(state.formulas.every((f) => f?.source === undefined)).toBe(true)
  })
})
