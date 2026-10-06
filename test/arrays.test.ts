/**
 * Arrays inside a cell beyond the Excel case set (`bench/coverage/sets/arrays.ts`): whole columns computed on
 * their occupied rows, laziness kept where nothing is an array, the code of a formula without arrays, results
 * that are arrays in a session, array constants in the canonical form.
 */
import { describe, expect, it } from 'vitest'
import { compute, debugGeneratedCode, openSession, restructure, ruRU, validateFormula, validateLocalFormula, type CellInput, type CellValue } from '../src/index.js'

const clock = { now: '2026-10-06T12:00:00Z', seed: 1 }

function sheet(cells: Record<string, CellInput>): Record<string, CellValue> {
  const r = compute({ sheets: [{ name: 'S', cells }] }, { clock })
  return Object.fromEntries(Object.entries(r.sheets[0].cells).map(([a, c]) => [a, c.value]))
}

describe('arrays beyond the Excel cases', () => {
  it('a whole column is 1,048,576 rows: the blank ones count where an operator makes them values', () => {
    const cells: Record<string, CellInput> = {}
    for (let r = 1; r <= 1000; r++) cells[`A${r}`] = { value: r % 3 === 0 ? 'x' : 'y' }
    for (let r = 1; r <= 1000; r++) cells[`B${r}`] = { value: r }
    const f = ['SUM(B:B*1)', 'SUM(B:B+1)', 'COUNT(B:B*1)', 'SUMPRODUCT((A:A="x")*B:B)', 'SUM(FILTER(B:B,A:A="x"))', 'AVERAGE(B:B*0)', 'SUMPRODUCT(B:B+1)', 'MATCH(1,(A:A="x")*(B:B>500),0)']
    f.forEach((x, i) => (cells[`D${i + 1}`] = { formula: x }))
    const v = sheet(cells)
    expect(f.map((_, i) => v[`D${i + 1}`])).toEqual([500500, 500500 + 1048576, 1048576, 166833, 166833, 0, 500500 + 1048576, 501])
  })

  it('IF and IFERROR stay lazy when their condition is one value, even with an array in a branch', () => {
    const v = sheet({ A1: { value: 1 }, A2: { value: 2 }, B1: { formula: 'IF(A1>0,SUM(A1:A2*2),1/0)' }, B2: { formula: 'IFERROR(A1,SUM(1/{0,1}))' }, B3: { formula: 'SUM(IF(A1:A2>1,A1:A2*10,1/0))' } })
    expect([v.B1, v.B2, v.B3]).toEqual([6, 1, { error: 'DIV0' }])
  })

  it('a formula without arrays compiles to the same scalar code; lifting appears only where an array can be', () => {
    const code = debugGeneratedCode({ sheets: [{ name: 'S', cells: { A1: { formula: 'B1*2+ROUND(C1,1)' }, A2: { formula: 'SUM(B1:B3*2)' } } }] }, { clock })
    expect(code).toContain('add(mul(cell(0,0,1),2),F.ROUND(cell(0,0,2),1))')
    expect(code).toContain('F.SUM(L2(mul,h0,2))')
  })

  it('a session moves a cell between one value and an array as its data change', () => {
    const s = openSession({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, B1: { value: 5 }, B2: { value: 6 }, C1: { formula: 'INDEX(B1:B2,A1)' }, C2: { formula: 'C1*2' } } }] }, { clock })
    expect(s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 0 } }])).toEqual([
      { sheet: 'S', cell: 'A1', value: 0 },
      { sheet: 'S', cell: 'C1', value: { error: 'SPILL' } },
      { sheet: 'S', cell: 'C2', value: { error: 'SPILL' } },
    ])
    expect(s.result().issues.find((i) => i.cell === 'C1')).toMatchObject({ check: 'unsupported', error: 'SPILL' })
    expect(s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 2 } }]).map((c) => [c.cell, c.value])).toEqual([['A1', 2], ['C1', 6], ['C2', 12]])
    expect(s.result().issues).toEqual([])
  })

  it('array constants: the canonical form, names and validation', () => {
    expect(validateFormula('sum({1,-2;"a",true})')).toMatchObject({ ok: true, canonical: 'SUM({1,-2;"a",TRUE})' })
    expect(validateFormula('{1,2;3}')).toMatchObject({ ok: false, code: 'syntax' })
    // typed in ru-RU: `;` between columns, `:` between rows
    expect(validateLocalFormula('СУММ({1,5;2:3;4})', ruRU)).toMatchObject({ ok: true, canonical: 'SUM({1.5,2;3,4})' })
    const r = compute({ sheets: [{ name: 'S', cells: { A1: { formula: 'SUM(x*{1;2})' }, B1: { value: 3 }, B2: { value: 4 } } }], names: [{ name: 'x', ref: "'S'!$B$1:$B$2" }] }, { clock })
    expect(r.sheets[0].cells.A1.value).toBe(11)
    // a formula with a constant parses, so a structure operation moves its references
    const moved = restructure({ sheets: [{ name: 'S', cells: { B1: { formula: 'SUM(A1:A2*{1;2})' } } }] }, { op: 'insertRows', sheet: 'S', row: 1 })
    expect(moved.ok && moved.book.sheets[0].cells.B2).toEqual({ formula: 'SUM(A2:A3*{1;2})' })
  })
})
