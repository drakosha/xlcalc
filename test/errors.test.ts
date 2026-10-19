/** Errors as values: the code in the cell, the reason in `diagnostics`; an invalid formula does not break the computation. */
import { describe, expect, it } from 'vitest'
import { compute, type BookInput, type CellInput } from '../src/index.js'
import { cellIssues } from './diagnostics.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 1 }
function run(cells: Record<string, CellInput>, extra: BookInput['sheets'] = []) {
  const r = compute({ sheets: [{ name: 'S', cells }, ...extra] }, { clock })
  const v = (a: string) => r.sheets[0].cells[a].value
  const issue = (a: string) => cellIssues(r).find((i) => i.sheet === 'S' && i.cell === a)
  return { r, v, issue }
}

describe('errors', () => {
  it('#DIV/0! through SUM and IFERROR, the reason names the source cell', () => {
    const { v, issue } = run({ A1: { value: 1 }, A2: { formula: '1/0' }, A3: { formula: 'SUM(A1:A2)' }, A4: { formula: 'IFERROR(SUM(A1:A2),-1)' }, A5: { formula: 'A3*2' } })
    expect(v('A2')).toEqual({ error: 'DIV0' })
    expect(v('A3')).toEqual({ error: 'DIV0' })
    expect(v('A4')).toBe(-1)
    expect(issue('A2')!.reason).toBe('division by zero')
    expect(issue('A3')!.reason).toBe("division by zero, from 'S'!A2")
    expect(issue('A5')!.reason).toBe("division by zero, from 'S'!A3")
  })

  it('#REF!: unknown sheet and literal', () => {
    const { v, issue } = run({ A1: { formula: 'Нет!A1+1' }, A2: { formula: '#REF!' } })
    expect(v('A1')).toEqual({ error: 'REF' })
    expect(issue('A1')).toMatchObject({ check: 'unknown_sheet', reason: 'unknown sheet Нет' })
    expect(v('A2')).toEqual({ error: 'REF' })
  })

  it('#NAME?: unknown function and name', () => {
    const { v, issue } = run({ A1: { formula: 'INDIRECT("B1")' }, A2: { formula: 'nope*2' } })
    expect(v('A1')).toEqual({ error: 'NAME' })
    expect(issue('A1')).toMatchObject({ check: 'unknown_function', at: 0 })
    expect(v('A2')).toEqual({ error: 'NAME' })
    expect(issue('A2')!.check).toBe('unknown_name')
  })

  it('cycle: the cells in the cycle and their dependents hold their values, 0 without held, each with its reason; the rest is computed', () => {
    const { v, issue } = run({ A1: { formula: 'A2+1', held: 4 }, A2: { formula: 'A1+1' }, A3: { formula: 'A1*2', held: 'x' }, A4: { formula: '5' }, B1: { formula: 'B1' } })
    expect([v('A1'), v('A2'), v('A3'), v('B1')]).toEqual([4, 0, 'x', 0])
    expect(v('A4')).toBe(5)
    for (const a of ['A1', 'A2', 'A3', 'B1']) expect(issue(a)).toMatchObject({ error: 'CYCLE', reason: 'circular reference: not calculated, the value is held' })
    expect(issue('A4')).toBeUndefined()
  })

  it('an array that cannot spill: #SPILL! with the cell in its way', () => {
    const { v, issue } = run({ A1: { value: 1 }, A2: { value: 2 }, B1: { formula: 'A1:A2*10' }, B2: { formula: 'SUM(A1:A2*10)' }, B3: { formula: 'B1+1' }, B4: { formula: 'INDEX(A1:A2*10,0,1)' } })
    expect([v('B1'), v('B2'), v('B3'), v('B4'), v('B5')]).toEqual([{ error: 'SPILL' }, 30, { error: 'SPILL' }, 10, 20])
    expect(issue('B1')).toEqual({ sheet: 'S', cell: 'B1', error: 'SPILL', reason: "the array of 2×1 values cannot spill: 'S'!B2 is not empty" })
    expect(issue('B3')).toMatchObject({ reason: "array result does not fit, from 'S'!B1" })
  })

  it('invalid formulas do not break the computation: a comma outside a function gives syntax', () => {
    const { v, issue, r } = run({ A1: { formula: 'SUM(B1:B2),SUM(C1:C2)' }, A2: { formula: '' }, A3: { formula: 'A1+1' }, A4: { formula: '2+2' } })
    expect(issue('A1')!.check).toBe('syntax')
    expect(issue('A2')!.check).toBe('empty')
    expect(v('A3')).toEqual({ error: 'ERROR' })
    expect(v('A4')).toBe(4)
    expect(cellIssues(r).map((i) => i.cell)).toEqual(['A1', 'A2', 'A3'])
  })

  it('external functions: the value comes from externals, a missing value is #N/A', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: 'BALANCE("2025-01")*2' }, A2: { formula: 'BALANCE("2025-02")' } } }] }
    const r = compute(book, { clock, externals: { functions: ['BALANCE'], values: { 'BALANCE("2025-01")': 21 } } })
    expect(r.sheets[0].cells.A1.value).toBe(42)
    expect(r.sheets[0].cells.A2.value).toEqual({ error: 'NA' })
  })
})
