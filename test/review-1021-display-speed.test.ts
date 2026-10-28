/**
 * The review of 21.10 on the speed of displayed text and typed entries: what each shortcut keeps as it was. Dates by
 * civil arithmetic against `Date`; widths by character code with a pair of surrogates as one character; a section's
 * plan read again for values of every kind; a typed formula's format with and without the book's links; the separators
 * of each locale's typed numbers; formulas shown in a locale kept by locale, links and table.
 */
import { describe, expect, it } from 'vitest'
import { daysInMonth, serialExcel, utcDays, ymd } from '../src/dates.js'
import { DAY_MS, EPOCH_MS, serialToUtcMs } from '../src/model.js'
import { createSession, displayValue, parseEntry, toLocal } from '../src/index.js'
import { showFormula } from '../src/locale/index.js'
import { ruRU } from '../src/locales/ru-RU.js'
import { advance, fitText, glyph, textWidth } from '../src/numfmt/glyphs.js'
import { mulberry32 } from '../src/synth.js'
import './locales.js'

/** The calendar by `Date`, as the functions read it before the arithmetic. */
function dateDays(y: number, m: number, d: number): number {
  const t = new Date(0)
  t.setUTCFullYear(y, m - 1, d)
  return Math.round((t.getTime() - EPOCH_MS) / DAY_MS)
}
function dateYmd(serial: number) {
  const s = Math.floor(serial)
  if (s === 0) return { y: 1900, m: 1, d: 0 }
  if (s === 60) return { y: 1900, m: 2, d: 29 }
  const t = new Date(serialToUtcMs(s))
  return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() }
}

describe('dates by civil arithmetic', () => {
  it('gives the days and dates Date gives, overflow, fractions, NaN and far dates included', () => {
    const rnd = mulberry32(21)
    for (let i = 0; i < 20_000; i++) {
      const y = Math.floor((rnd() - 0.5) * 30_000)
      const m = Math.floor((rnd() - 0.5) * 100)
      const d = Math.floor((rnd() - 0.5) * 2_000)
      expect(utcDays(y, m, d)).toBe(dateDays(y, m, d))
      const s = Math.floor((rnd() - 0.2) * 4e6) + rnd()
      expect(ymd(s)).toEqual(dateYmd(s))
    }
    const edges: [number, number, number][] = [
      [1900, 1, 0],
      [1900, 2, 29],
      [1900, 0.5, 1],
      [2024, -0.5, 15],
      [2024, 13.9, 1.9],
      [-1, 1, 1],
      [-400, 3, 1],
      [9999, 12, 31],
      [10_000, 1, 1],
      [100_000, 100_000, 10_000_000],
      [100_001, 1, 1],
      [300_000, 1, 1],
      [2024, NaN, 1],
      [Infinity, 1, 1],
      [2024, 1, 32_768],
    ]
    for (const [y, m, d] of edges) expect(utcDays(y, m, d)).toBe(dateDays(y, m, d))
    for (const s of [0, 0.5, 59, 60, 60.9, 61, -1, -693_960, 2_958_465, 9e7, 1e8, 1e9, NaN, Infinity]) expect(ymd(s)).toEqual(dateYmd(s))
    expect([serialExcel(1900, 2, 29), serialExcel(1900, 3, 1), serialExcel(2024, 2, 29)]).toEqual([60, 61, 45351])
  })

  it('counts the days of a month, February of a leap year too, and none for a month that is not one', () => {
    expect([1900, 2000, 2023, 2024].map((y) => daysInMonth(y, 2))).toEqual([28, 29, 28, 29])
    expect([1, 4, 12].map((m) => daysInMonth(2023, m))).toEqual([31, 30, 31])
    expect([daysInMonth(2023, 0), daysInMonth(2023, 13)]).toEqual([undefined, undefined])
  })
})

describe('widths by character code', () => {
  it('measures a pair of surrogates as one character, a digit wide, and a lone one as a digit', () => {
    expect([glyph('a'), glyph('b'), glyph('Щ'), glyph('₽'), glyph('😀'), glyph(''), glyph('ab')]).toEqual([5, 6, 10, 6, 6, 6, 6])
    expect([textWidth('a😀b'), textWidth('\ud800x'), textWidth('xy\udc00'), textWidth('Щука₽'), textWidth('WWW'), textWidth('')]).toEqual([17, 11, 16, 31, 30, 0])
    expect([advance('0'), advance('😀'), advance(' ')]).toEqual([1038, 1038, advance(' ')])
    expect(fitText('é😀')).toBe(fitText('é0'))
  })

  it('shows text under a text section whole, without a fill whatever its width, with one filling the cell', () => {
    expect(['a😀b', 'Оплата', 'x'.repeat(300)].map((t) => displayValue(t, { numberFormat: '@', width: 1 }).text.length)).toEqual([4, 6, 300])
    expect([3, 5].map((width) => displayValue('ab', { numberFormat: '@*-', width }).text)).toEqual(['ab', 'ab----'])
    expect([1, 4].map((width) => displayValue(true, { numberFormat: '@', width }).text)).toEqual(['', '###'])
  })
})

describe("a section's plan, read again for each value", () => {
  it('shows General in a narrow column as it did, about the exponents of two and three digits', () => {
    const cases: [number, string[]][] = [
      [9.99999e99, ['#', '##', '###', '#####', '######', '1E+100']],
      [9.999995e99, ['#', '##', '###', '#####', '######', '1E+100']],
      [-1.23456e100, ['#', '##', '###', '#####', '######', '-1E+100']],
      [9.99995e-100, ['0', '0', '0', '0', '1E-99', '1E-99']],
      [1.5e-101, ['0', '0', '0', '0', '######', '2E-101']],
      [1.23456789e21, ['#', '##', '###', '#####', '1E+21', '1.2E+21']],
      [-1.23456789e-21, ['#', '-0', '-0', '-0', '-0', '-1.2E-21']],
      [123456789012, ['#', '##', '###', '#####', '1E+11', '1.2E+11']],
      [0.000123456789, ['0', '0', '0', '0', '1E-04', '0.00012']],
    ]
    for (const [v, texts] of cases) expect([2, 3, 4, 5, 6, 8.43].map((width) => displayValue(v, { width }).text)).toEqual(texts)
  })

  it('shows zero and other numbers in turn by scientific notation, a fraction, a percent and a scale', () => {
    const cases: [string, string[]][] = [
      ['0.00E+00', ['0.00E+00', '1.23E+04', '0.00E+00', '-5.00E-01', '1.50E+00', '0.00E+00']],
      ['##0.0E+0', ['000.0E+0', '12.3E+3', '000.0E+0', '-500.0E-3', '1.5E+0', '000.0E+0']],
      ['# ?/?', ['0    ', '12345 2/3', '0    ', '- 1/2', '1 1/2', '0    ']],
      ['# ??/16', ['0      ', '12345 11/16', '0      ', '-  8/16', '1  8/16', '0      ']],
      ['0.0%', ['0.0%', '1234567.8%', '0.0%', '-50.0%', '150.0%', '0.0%']],
      ['#,##0.00,,"M"', ['0.00M', '0.01M', '0.00M', '0.00M', '0.00M', '0.00M']],
    ]
    for (const [numberFormat, texts] of cases) expect([0, 12345.678, 0, -0.5, 1.5, 0].map((v) => displayValue(v, { numberFormat }).text)).toEqual(texts)
  })

  it('chooses conditional sections and their minus for values in any order', () => {
    const shown = (numberFormat: string) => [-5, 5, 0, -7, 1234567, 1234].map((v) => displayValue(v, { numberFormat }))
    expect([-5, 5, 0, -7].map((v) => displayValue(v, { numberFormat: '[<0]0;@', width: 4 }).text)).toEqual(['-5', '###', '###', '-7'])
    expect(shown('[Red][<0]0.00;[Blue][>=0]0.00')).toEqual([
      { text: '5.00', color: '#FF0000' },
      { text: '5.00', color: '#0000FF' },
      { text: '0.00', color: '#0000FF' },
      { text: '7.00', color: '#FF0000' },
      { text: '1234567.00', color: '#0000FF' },
      { text: '1234.00', color: '#0000FF' },
    ])
    expect(shown('[>=1000000]0.0,,"M";[>=1000]0.0,"K";0').map((d) => d.text)).toEqual(['-5', '5', '0', '-7', '1.2M', '1.2K'])
    expect(shown('[<0]0;[<-5]0;0').map((d) => d.text)).toEqual(['5', '5', '0', '7', '1234567', '1234'])
  })

  it('shows dates and times by a section read once: minutes and months, decimals of a second, AM/PM, the calendar\'s end', () => {
    const cases: [string, string[]][] = [
      ['mm:ss', ['00:00', '00:00', '59:59', '00:00', '57:47']],
      ['mm/dd/yyyy', ['03/15/2023', '01/00/1900', '12/31/9999', '01/01/1904', '03/01/1900']],
      ['h:mm AM/PM', ['6:00 PM', '12:00 PM', '11:59 PM', '6:00 AM', '2:57 AM']],
      ['[h]:mm:ss.00', ['1080018:00:00.00', '12:00:00.00', '71003183:59:59.14', '35094:00:00.00', '1466:57:46.60']],
      ['dd.mm.yyyy hh:mm', ['15.03.2023 18:00', '00.01.1900 12:00', '31.12.9999 23:59', '01.01.1904 06:00', '01.03.1900 02:57']],
      ['yy.00.00', ['23.00.00', '00.00.00', '99.14.14', '04.00.00', '00.60.60']],
    ]
    for (const [numberFormat, texts] of cases) expect([45000.75, 0.5, 2958465.99999, 1462.25, 61.123456].map((v) => displayValue(v, { numberFormat }).text)).toEqual(texts)
    expect(displayValue(2958466, { numberFormat: 'mm/dd/yyyy', width: 8.43 }).text).toBe('#########')
  })
})

describe('typed entries', () => {
  const links = [{ target: 'Book.xlsx', sheets: [{ name: 'Data' }] }]

  it("takes a formula's format from its one check, and none where the book's links read a link", () => {
    expect(parseEntry('=TODAY()+1')).toEqual({ ok: true, input: { formula: 'TODAY()+1', numberFormat: 14 } })
    expect(parseEntry('=TODAY()+1', { externalBooks: links })).toEqual({ ok: true, input: { formula: 'TODAY()+1', numberFormat: 14 } })
    expect(parseEntry('=СЕГОДНЯ()+1', { locale: 'ru-RU' })).toEqual({ ok: true, input: { formula: 'TODAY()+1', numberFormat: 14 } })
    expect(parseEntry('=TODAY()+[Book.xlsx]Data!A1', { externalBooks: links })).toEqual({ ok: true, input: { formula: "TODAY()+'[1]Data'!A1" } })
    expect(parseEntry('=ДАТА(2020;1;1)+[Book.xlsx]Data!A1', { locale: 'ru-RU', externalBooks: links })).toEqual({ ok: true, input: { formula: "DATE(2020,1,1)+'[1]Data'!A1" } })
  })

  it("reads each locale's separators in numbers typed in turn", () => {
    const typed = (text: string, locale: string) => parseEntry(text, { locale })
    for (let i = 0; i < 2; i++) {
      expect(typed('1,5', 'ru-RU')).toEqual({ ok: true, input: { value: 1.5 } })
      expect(typed('1,5', 'en-US')).toEqual({ ok: true, input: { value: '1,5' } })
      expect(typed('1,5E+3', 'ru-RU')).toEqual({ ok: true, input: { value: 1500, numberFormat: 11 } })
      expect(typed('1E+400', 'en-US')).toEqual({ ok: true, input: { value: '1E+400', numberFormat: 11 } })
      expect(typed('1,5E+400', 'ru-RU')).toEqual({ ok: true, input: { value: '1,5E+400', numberFormat: 11 } })
      expect(typed('1 234,5', 'ru-RU')).toEqual({ ok: true, input: { value: 1234.5, numberFormat: 4 } })
      expect(typed('1,234.5', 'en-US')).toEqual({ ok: true, input: { value: 1234.5, numberFormat: 4 } })
    }
  })
})

describe('formulas shown in a locale, kept between calls', () => {
  it('by the list of links, the table and the locale', () => {
    const a = [{ target: 'Book.xlsx', sheets: [{ name: 'Data' }] }]
    const b = [{ target: 'Other.xlsx', sheets: [{ name: 'Data' }] }]
    const f = "SUM('[1]Data'!A1,1.5)"
    expect([a, b, undefined, a].map((x) => showFormula(f, ruRU, undefined, x))).toEqual([
      { ok: true, formula: 'СУММ([Book.xlsx]Data!A1;1,5)' },
      { ok: true, formula: 'СУММ([Other.xlsx]Data!A1;1,5)' },
      { ok: true, formula: "СУММ('[1]Data'!A1;1,5)" },
      { ok: true, formula: 'СУММ([Book.xlsx]Data!A1;1,5)' },
    ])
    expect([showFormula('SUM(Sales[Qty])', ruRU, 'Sales'), showFormula('SUM(Sales[Qty])', ruRU), showFormula('SUM(Sales[Qty])', 'en-US', 'Sales')]).toEqual([
      { ok: true, formula: 'СУММ([Qty])' },
      { ok: true, formula: 'СУММ(Sales[Qty])' },
      { ok: true, formula: 'SUM([Qty])' },
    ])
    expect(toLocal('SUM(Sales[Qty])', ruRU, { table: 'Sales' })).toEqual({ ok: true, formula: 'СУММ([Qty])' })
  })

  it("in a window shown again, after an edit of a cell's formula", () => {
    const s = createSession({ locale: 'ru-RU', sheets: [{ name: 'S', cells: { A1: { value: 1.5 }, B1: { formula: 'SUM(A1,2.5)' } } }] })
    const shown = () => s.getViewport('S', 'B1', { formulaLocal: true, editText: true }).cells.B1
    expect([shown().formulaLocal, shown().editText]).toEqual(['СУММ(A1;2,5)', '=СУММ(A1;2,5)'])
    s.setCells([{ sheet: 'S', cell: 'B1', input: { formula: 'ROUND(A1,0)' } }])
    expect([shown().formulaLocal, shown().editText]).toEqual(['ОКРУГЛ(A1;0)', '=ОКРУГЛ(A1;0)'])
  })
})
