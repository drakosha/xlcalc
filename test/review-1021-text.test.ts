/**
 * Reproductions from the review of 21.10 on text and numbers: number text padded with a tab, a line break or a
 * non-breaking space is no number to operators and numeric arguments, as it is none to VALUE (Excel trims spaces only).
 * With closures and eval.
 */
import { describe, expect, it } from 'vitest'
import { compute, type BookInput, type CellInput } from '../src/index.js'
import './locales.js'

/** The values of the formulas, one per cell of column B, with `cells` beside them. */
function values(formulas: string[], execution: 'closures' | 'eval', locale?: string, cells: Record<string, CellInput> = {}) {
  const book: BookInput = { ...(locale ? { locale } : {}), sheets: [{ name: 'S', cells: { ...cells, ...Object.fromEntries(formulas.map((f, i) => ['B' + (i + 1), { formula: f }])) } }] }
  const r = compute(book, { execution })
  return formulas.map((_, i) => r.sheets[0].cells['B' + (i + 1)].value)
}

describe.each(['closures', 'eval'] as const)('whitespace around number text, %s', (execution) => {
  const value = { error: 'VALUE' }

  it('refuses number text with a tab, a line break or a non-breaking space around it', () => {
    const pads = ['CHAR(9)&"5"', '"5"&CHAR(9)', 'CHAR(10)&"5"', 'CHAR(13)&"5"', 'UNICHAR(160)&"5"', 'UNICHAR(8195)&"5"', 'CHAR(9)&"5.5"', 'CHAR(9)&"1E+3"']
    const forms = pads.flatMap((p) => [`(${p})+0`, `--(${p})`, `ABS(${p})`, `SUM(${p})`, `VALUE(${p})`])
    expect(values(forms, execution)).toEqual(forms.map(() => value))
    const cells = { A1: { value: '\t5' }, A2: { value: ' 5' } }
    expect(values(['A1+0', 'A1*1', 'INT(A1)', 'A2+0', 'VALUE(A2)'], execution, undefined, cells)).toEqual([value, value, value, value, value])
  })

  it('still trims spaces', () => {
    expect(values(['" 5"+0', '"5 "+0', 'ABS("  5.5 ")', 'VALUE(" 5")'], execution)).toEqual([5, 5, 5.5, 5])
  })

  it('agrees with ru-RU, which already refused them', () => {
    expect(values(['(CHAR(9)&"5")+0', '(UNICHAR(160)&"5")+0', '" 5"+0'], execution, 'ru-RU')).toEqual([value, value, 5])
  })
})
