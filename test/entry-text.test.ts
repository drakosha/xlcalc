/**
 * `entryText` and `needsApostrophe`: the text a cell opens with for editing, in en-US and ru-RU, read back by `parseEntry`
 * to the same value (numbers to 15 digits, dates and times to the second), and the apostrophe exactly where text would
 * read as something else.
 */
import { describe, expect, it } from 'vitest'
import type { CellValue, NumberFormat } from '../src/contract.js'
import { entryText, InputError, needsApostrophe, parseEntry, sameValue } from '../src/index.js'
import './locales.js'

const NOW = '2026-10-10T12:00:00Z'

describe('entryText', () => {
  it('writes a value as the formula bar does', () => {
    const en = (v: CellValue, numberFormat?: NumberFormat) => entryText(v, { ...(numberFormat === undefined ? {} : { numberFormat }) })
    const ru = (v: CellValue, numberFormat?: NumberFormat) => entryText(v, { locale: 'ru-RU', ...(numberFormat === undefined ? {} : { numberFormat }) })
    expect([en(1234.5), ru(1234.5), en(-0.5, '$#,##0.00'), en(1 / 3, '# ?/?'), en(1e20), en(1e21), en(1e-7), en(1e-20)]).toEqual(['1234.5', '1234,5', '-0.5', '0.333333333333333', '100000000000000000000', '1E+21', '0.0000001', '1E-20'])
    expect([en(0.12, '0%'), en(0.123456, '0.00%'), ru(0.125, 10), en(-0.05, 9)]).toEqual(['12%', '12.3456%', '12,5%', '-5%'])
    expect([en(45688, 14), ru(45688, 'd mmmm yyyy'), en(45688.75, 22), ru(45688.75, 22)]).toEqual(['1/31/2025', '31.01.2025', '1/31/2025  6:00:00 PM', '31.01.2025  18:00:00'])
    expect([en(0.75, 'h:mm'), ru(0.75, 'h:mm'), en(1.5, '[h]:mm:ss'), en(0, 14), en(-1, 14), en(2958466, 14)]).toEqual(['6:00:00 PM', '18:00:00', '1/1/1900  12:00:00 PM', '1/0/1900', '-1', '2958466'])
    expect([en(true), ru(false), en({ error: 'DIV0' }), ru({ error: 'NA' }), en(null)]).toEqual(['TRUE', 'ЛОЖЬ', '#DIV/0!', '#Н/Д', ''])
    expect([en('007'), en('abc'), en('=A1'), en("'x"), ru('1.5'), en('1.5'), ru('1,5'), en('5', '0%'), en('007', '@'), en('')]).toEqual(["'007", 'abc', "'=A1", "''x", "'1.5", "'1.5", "'1,5", "'5", '007', "'"])
    expect(entryText(0, { numberFormat: 14, dateSystem: 1904 })).toBe('1/1/1904')
  })

  describe.each(['en-US', 'ru-RU'])('reads back to the value (%s)', (locale) => {
    const values: [CellValue, NumberFormat | undefined][] = [
      ...[0, 5, -5, 1234.5, -0.5, 1 / 3, 0.1 + 0.2, 1e20, 1e21, 1.5e-7, 1e-20, 123456789012345678, 3.14159265358979].map((n): [CellValue, NumberFormat | undefined] => [n, undefined]),
      [1234.5, '$#,##0.00'], [0.5, '# ?/?'], [1500, '0.00E+00'], [1234567.891, '#,##0.00'],
      [0.12, '0%'], [0.123456, '0.00%'], [-0.05, 9], [12.34, 9],
      [45688, 14], [45688.75, 22], [45688.770833333336, 'd mmmm yyyy h:mm'], [45688.5, 'mmm-yy'], [2958465, 14], [-1, 14],
      [0.75, 'h:mm'], [0.7710069444444444, 21], [1.0625, '[h]:mm:ss'], [1.75, 'h:mm'],
      [true, undefined], [false, undefined],
      ...['007', 'abc', '=1+2', "'x", '1.5', '1,5', 'TRUE', 'ИСТИНА', '12%', '1/31/2025', '31.01.2025', '#N/A', ' 12 ', '-C1', '+5', ''].map((t): [CellValue, NumberFormat | undefined] => [t, undefined]),
      ['5', '0%'], ['007', '@'], ['abc', 14],
    ]
    it.each(values.map(([v, f]) => [JSON.stringify(v), JSON.stringify(f ?? null), v, f] as const))('%s by %s', (_, __, v, numberFormat) => {
      const text = entryText(v, { locale, ...(numberFormat === undefined ? {} : { numberFormat }) })
      // a percent cell's editor adds `%` to a number typed into it: the bar's text reads back in General
      const cell = numberFormat === undefined || typeof v === 'number' ? undefined : { numberFormat }
      const r = parseEntry(text, { locale, now: NOW, ...(cell ? { cell } : {}) })
      if (!r.ok || !r.input) throw new Error(`${text} does not read back`)
      const back = r.input.value ?? null
      expect(typeof v === 'number' ? sameValue(back as CellValue, v) : back, text).toBe(typeof v === 'number' ? true : v)
    })
  })

  it('an error reads back as the formula of that error, as parseEntry takes one typed', () => {
    expect(parseEntry(entryText({ error: 'NA' }, { locale: 'ru-RU' }), { locale: 'ru-RU' })).toEqual({ ok: true, input: { formula: '#N/A' } })
  })
})

describe('needsApostrophe', () => {
  it('where parseEntry would read the text as something else', () => {
    const cases: [string, boolean, string?, NumberFormat?][] = [
      ['007', true], ['abc', false], ['1.5', true], ['1.5', true, 'ru-RU'], ['1,5', false], ['1,5', true, 'ru-RU'], ['TRUE', true], ['ИСТИНА', false], ['ИСТИНА', true, 'ru-RU'],
      ['=A1', true], ["'x", true], ['#N/A', true], ['', true], ['5', true, undefined, '0%'], ['5', false, undefined, '@'], ['abc', false, undefined, '0%'], ['-abc', true], ['15.3', true, 'ru-RU'],
    ]
    for (const [text, want, locale, numberFormat] of cases) expect(needsApostrophe(text, { ...(locale ? { locale } : {}), ...(numberFormat === undefined ? {} : { numberFormat }), now: NOW }), `${text} ${locale ?? ''} ${numberFormat ?? ''}`).toBe(want)
  })

  it('wrong input is an InputError', () => {
    expect(() => needsApostrophe(5 as unknown as string)).toThrow(InputError)
    expect(() => entryText(Number.NaN)).toThrow(InputError)
    expect(() => entryText(1, { format: '0%' } as never)).toThrow(/unknown field "format"/)
    expect(() => entryText(1, { numberFormat: 99 })).toThrow(/built-in format/)
    expect(() => entryText(1, { locale: 'xx' })).toThrow(/unknown locale/)
  })
})
