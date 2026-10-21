/**
 * Reproductions from the review of 21.10 on formula text: a tab or a line break between a function's name and its `(`
 * keeps the call, as a space does, so the canonical text means what the formula meant. With closures and eval.
 */
import { describe, expect, it } from 'vitest'
import { compute, validateFormula, validateLocalFormula } from '../src/index.js'
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
