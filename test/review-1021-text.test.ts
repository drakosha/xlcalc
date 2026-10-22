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

describe.each(['closures', 'eval'] as const)('FIXED and DOLLAR of large numbers, %s', (execution) => {
  it('write the digits from 1E+21 on, 15 significant ones and zeros, as TEXT does', () => {
    const cases: [string, string][] = [
      ['FIXED(1E21)', '1,000,000,000,000,000,000,000.00'],
      ['FIXED(1E21,2,TRUE)', '1000000000000000000000.00'],
      ['FIXED(-1E21,0)', '-1,000,000,000,000,000,000,000'],
      ['FIXED(1E21,-5)', '1,000,000,000,000,000,000,000'],
      ['FIXED(123456789012345678901234,0)', '123,456,789,012,345,000,000,000'],
      ['DOLLAR(1E21)', '$1,000,000,000,000,000,000,000.00'],
      ['DOLLAR(-1E21)', '($1,000,000,000,000,000,000,000.00)'],
      ['FIXED(2^60,0)', '1,152,921,504,606,850,000'],
      ['FIXED(1/3,20)', '0.33333333333333300000'],
      ['FIXED(1.005,2)', '1.01'],
      ['FIXED(1234.567,-2)', '1,200'],
      ['FIXED(1E20)', '100,000,000,000,000,000,000.00'],
    ]
    expect(values(cases.map((c) => c[0]), execution)).toEqual(cases.map((c) => c[1]))
    expect(values(['FIXED(1E21,1)'], execution, 'ru-RU')).toEqual(['1 000 000 000 000 000 000 000,0'])
  })
})

describe.each(['closures', 'eval'] as const)('TEXTBEFORE, TEXTAFTER and TEXTSPLIT in any case past an İ, %s', (execution) => {
  it('cut the text where the delimiter is, as in their own case', () => {
    const cases: [string, unknown][] = [
      ['TEXTBEFORE("İstanbul","s",1,1)', 'İ'],
      ['TEXTAFTER("İstanbul","s",1,1)', 'tanbul'],
      ['TEXTAFTER("İstanbul","S",-1,1)', 'tanbul'],
      ['TEXTBEFORE("İİİ;x",";",1,1)', 'İİİ'],
      ['TEXTAFTER("aİbİc","B",1,1)', 'İc'],
      ['TEXTBEFORE("x;İ",";",1,1)', 'x'],
      ['TEXTBEFORE("İxAy","a",1,1)', 'İx'],
      ['INDEX(TEXTSPLIT("İ-s-t","S",,,1),1)', 'İ-'],
      ['INDEX(TEXTSPLIT("İ-s-t","S",,,1),2)', '-t'],
      ['ROWS(TEXTSPLIT("İaXbxc",,"x",,1))', 3],
      ['INDEX(TEXTSPLIT("İaXbxc",,"x",,1),2,1)', 'b'],
    ]
    expect(values(cases.map((c) => c[0]), execution)).toEqual(cases.map((c) => c[1]))
  })
})
