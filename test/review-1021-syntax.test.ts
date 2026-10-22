/**
 * Reproductions from the review of 21.10 on formula text: a tab or a line break between a function's name and its `(`
 * keeps the call, as a space does, so the canonical text means what the formula meant. With closures and eval.
 */
import { describe, expect, it } from 'vitest'
import { compute, createSession, parseEntry, validateFormula, validateLocalFormula } from '../src/index.js'
import './locales.js'

describe.each(['closures', 'eval'] as const)('whitespace before a call, %s', (execution) => {
  const calls = ['SUM\t(A1)', 'SUM\n(A1)', 'SUM\r\n(A1)', 'SUM \n (A1)', 'ROUND\n(A1,0)']

  it('reads a function name followed by a tab or a line break and `(` as a call', () => {
    const cells = Object.fromEntries(calls.map((f, i) => ['B' + (i + 1), { formula: f }]))
    const r = compute({ sheets: [{ name: 'S', cells: { A1: { value: 5 }, ...cells } }] }, { execution })
    expect(calls.map((_, i) => r.sheets[0].cells['B' + (i + 1)].value)).toEqual([5, 5, 5, 5, 5])
  })

  it('keeps the canonical text a call that reads back the same', () => {
    for (const f of calls) {
      const v = validateFormula(f)
      expect(v.ok && { canonical: v.canonical, names: v.names }).toEqual({ canonical: f.startsWith('SUM') ? 'SUM(A1)' : 'ROUND(A1,0)', names: [] })
    }
    expect(validateLocalFormula('СУММ\t(A1)', 'ru-RU')).toMatchObject({ ok: true, canonical: 'SUM(A1)' })
  })

  it('keeps a line break between two references an intersection', () => {
    const r = compute({ sheets: [{ name: 'S', cells: { A1: { value: 5 }, B1: { formula: 'A1:A2\n(A1)' } } }] }, { execution })
    expect(r.sheets[0].cells.B1.value).toBe(5)
    expect(validateFormula('A1:A2\n(A1)')).toMatchObject({ ok: true, canonical: 'A1:A2 (A1)' })
  })
})

describe.each(['closures', 'eval'] as const)('a range typed with its corners reversed, %s', (execution) => {
  // each axis ordered on its own, the `$` going with its bound, as Excel shows a typed reference (docs/addresses.md)
  const typed: [string, string][] = [
    ['SUM(B2:A1)', 'SUM(A1:B2)'],
    ['SUM($B$2:A1)', 'SUM(A1:$B$2)'],
    ['SUM(B1:A2)', 'SUM(A1:B2)'],
    ['SUM($B1:A$2)', 'SUM(A1:$B$2)'],
    ['SUM(b:$a)', 'SUM($A:B)'],
    ['SUM(2:$1)', 'SUM($1:2)'],
    ['SUM(S!B2:A1)', "SUM('S'!A1:B2)"],
    ['SUM(A$1:$B2)', 'SUM(A$1:$B2)'],
  ]

  it('is stored with its corners in order, as a structure operation writes it', () => {
    const cells = Object.fromEntries(typed.map(([f], i) => ['D' + (i + 1), { formula: f }]))
    const r = compute({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, B2: { value: 2 }, ...cells } }] }, { execution })
    expect(typed.map((_, i) => r.sheets[0].cells['D' + (i + 1)].formula)).toEqual(typed.map(([, f]) => f))
    expect(validateFormula('SUM(B2:A1)')).toMatchObject({ ok: true, canonical: 'SUM(A1:B2)' })
    expect(parseEntry('=sum(b:a)')).toEqual({ ok: true, input: { formula: 'SUM(A:B)' } })
  })

  it('keeps one text through a structure operation', () => {
    const s = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, C1: { formula: 'SUM(B2:A1)' }, C2: { formula: 'SUM(B:A)' }, C3: { formula: 'SUM(2:1)' } } }] }, { execution })
    const formulas = (range: string) => Object.fromEntries(Object.entries(s.getCells('S', range)).map(([k, c]) => [k, c.formula]))
    expect(formulas('C1:C3')).toEqual({ C1: 'SUM(A1:B2)', C2: 'SUM(A:B)', C3: 'SUM(1:2)' })
    expect(s.editStructure({ op: 'insertRows', sheet: 'S', row: 2 }).ok).toBe(true)
    expect(formulas('C1:C4')).toEqual({ C1: 'SUM(A1:B3)', C3: 'SUM(A:B)', C4: 'SUM(1:3)' })
    s.close()
  })
})
