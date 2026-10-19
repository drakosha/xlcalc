/**
 * `parseEntry`'s contract: what it returns for each kind of typed text, the cell's style, meta and format, the clock's
 * year and the date system, positions of a formula Excel does not take, input checks, and the way into a session.
 * Excel's own answers are checked in `entry-excel.test.ts` and `entry-win.test.ts`.
 */
import { describe, expect, it } from 'vitest'
import { displayValue, InputError, createSession, parseEntry, type EntryResult } from '../src/index.js'
import './locales.js'

const now = '2026-10-10T12:00:00Z'
const ru = (text: string, cell?: object) => parseEntry(text, { locale: 'ru-RU', now, ...(cell ? { cell } : {}) })
const en = (text: string, cell?: object) => parseEntry(text, { now, ...(cell ? { cell } : {}) })
const input = (r: EntryResult) => (r.ok ? r.input : r)

describe('parseEntry', () => {
  it('reads numbers, dates, percent, booleans, errors and text by the locale, with the formats Excel sets', () => {
    expect(['1,5', '1 234,5', '15.03.2026', '5%', 'ИСТИНА', 'истина', 'TRUE', '#Н/Д', '#N/A', "'007", 'abc'].map((t) => input(ru(t)))).toEqual([
      { value: 1.5 },
      { value: 1234.5, numberFormat: 4 },
      { value: 46096, numberFormat: 14 },
      { value: 0.05, numberFormat: 9 },
      { value: true },
      { value: true },
      { value: 'TRUE' },
      { formula: '#N/A' },
      { value: '#N/A' },
      { value: '007' },
      { value: 'abc' },
    ])
    expect(['3/15/2026', '1,234.5', '1/2', '1 1/2', '1.5E+3', '6:30 PM', '25:00', '3/15/2026 18:30'].map((t) => input(en(t)))).toEqual([
      { value: 46096, numberFormat: 14 },
      { value: 1234.5, numberFormat: 4 },
      { value: 46024, numberFormat: 16 },
      { value: 1.5, numberFormat: 12 },
      { value: 1500, numberFormat: 11 },
      { value: 0.7708333333333334, numberFormat: 18 },
      { value: 1.0416666666666667, numberFormat: 46 },
      { value: 46096.770833333336, numberFormat: 22 },
    ])
  })

  it("gives typed money the locale's money format with red negatives", () => {
    const usd = input(en('$1,234.50')) as { value: number; numberFormat: string }
    const rub = input(ru('-1 234,5 ₽')) as { value: number; numberFormat: string }
    expect([usd.value, displayValue(-usd.value, { numberFormat: usd.numberFormat, width: 20 })]).toEqual([1234.5, { text: '($1,234.50)', color: '#FF0000' }])
    expect([rub.value, displayValue(rub.value, { numberFormat: rub.numberFormat, locale: 'ru-RU', width: 20 }).text]).toEqual([-1234.5, '-1 234,50 ₽'])
  })

  it('reads numbers with 15 significant digits and keeps out-of-range scientific text as text', () => {
    expect(['1234567890123456789', '1E+307', '1E+308', '1E-309'].map((t) => input(en(t)))).toEqual([{ value: 1234567890123450000 }, { value: 1e307, numberFormat: 11 }, { value: '1E+308', numberFormat: 11 }, { value: 0, numberFormat: 11 }])
  })

  it('takes a formula after =, and after + or - before what is no number; only a formula that does not parse stops the entry', () => {
    expect(['=СУММ(A1;1,5)', '=сумм(a1;1,5)', '=FOO(1)', '-5', '+5', '-C1', '+1+2', '-', '='].map((t) => input(ru(t)))).toEqual([
      { formula: 'SUM(A1,1.5)' },
      { formula: 'SUM(A1,1.5)' },
      { formula: 'FOO(1)' },
      { value: -5 },
      { value: 5 },
      { formula: '-C1' },
      { formula: '1+2' },
      { value: '-' },
      { value: '=' },
    ])
    expect(ru('=SUM(A1,1.5)')).toEqual({ ok: false, code: 'syntax', message: 'Здесь не может стоять «.»', at: 9 })
    expect(en('=1+')).toEqual({ ok: false, code: 'syntax', message: 'unexpected end of formula', at: 3 })
    expect(en('-1,5')).toMatchObject({ ok: false, code: 'syntax', at: 2 })
    expect(en('+1,5')).toMatchObject({ ok: false, code: 'syntax', at: 2 })
  })

  it('gives a formula in General the format of what it computes', () => {
    expect(['=TODAY()', '=NOW()', '=TIME(18,30,0)', '=TODAY()+1', '=YEAR(TODAY())', '=PMT(0.1,12,1000)'].map((t) => (input(en(t)) as { numberFormat?: unknown }).numberFormat)).toEqual([
      14,
      22,
      18,
      14,
      undefined,
      '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
    ])
  })

  it("keeps the cell's style and meta, and its format but where Excel changes it", () => {
    const cell = { numberFormat: '0.00', style: 2, meta: { id: 7 } }
    expect(input(ru('15.03.2026', cell))).toEqual({ value: 46096, numberFormat: '0.00', style: 2, meta: { id: 7 } })
    expect(input(ru('18:30', { numberFormat: 14, style: 1 }))).toEqual({ value: 0.7708333333333334, numberFormat: 20, style: 1 })
    expect(input(ru('15.03.2026', { numberFormat: 'd mmmm yyyy' }))).toEqual({ value: 46096, numberFormat: 'd mmmm yyyy' })
    expect(input(en('=1+2', { numberFormat: '@' }))).toEqual({ value: '=1+2', numberFormat: '@' })
    expect(input(en('5', { numberFormat: '0%' }))).toEqual({ value: 0.05, numberFormat: '0%' })
    expect(input(en('0.5', { numberFormat: '0%' }))).toEqual({ value: 0.005, numberFormat: '0%' })
    expect(input(en('$5', { numberFormat: '0%' }))).toEqual({ value: 5, numberFormat: '"$"#,##0_);[Red]\\("$"#,##0\\)' })
  })

  it("keeps the cell's note and link, as Excel keeps them whatever is typed", () => {
    const note = { author: 'Ann', text: 'Checked' }
    const link = { url: 'https://example.com/', tooltip: 'Open' }
    expect(input(en('5', { note, link, style: 1 }))).toEqual({ value: 5, style: 1, note, link })
    expect(input(en('=SUM(1,2)', { link }))).toEqual({ formula: 'SUM(1,2)', link })
    expect(en('', { note })).toEqual({ ok: true, input: { note } })
    expect(() => parseEntry('1', { cell: { note: { author: 'Ann' } as never } })).toThrow(new InputError('opts.cell.note.text', 'expected a string, got undefined'))
    expect(() => parseEntry('1', { cell: { link: { tooltip: 'x' } } })).toThrow(new InputError('opts.cell.link', 'a hyperlink needs a url, a location or both'))
  })

  it('clears the value of empty text and keeps the rest of the cell, null when nothing is left', () => {
    expect(en('')).toEqual({ ok: true, input: null })
    expect(en('', { numberFormat: '0.00', meta: 1 })).toEqual({ ok: true, input: { numberFormat: '0.00', meta: 1 } })
  })

  it("completes a date without a year by the year of `now`, and counts serials in the book's date system", () => {
    expect(input(parseEntry('15.3', { locale: 'ru-RU', now: '2025-06-01' }))).toEqual({ value: 45731, numberFormat: 16 })
    expect(input(parseEntry('3/15', { now: '2026-01-01T00:00:00Z' }))).toEqual({ value: 46096, numberFormat: 16 })
    expect(input(parseEntry('15.03.2026', { locale: 'ru-RU', dateSystem: 1904 }))).toEqual({ value: 44634, numberFormat: 14 })
    expect((input(parseEntry('1.5')) as { value: number }).value).toBe(1.5)
  })

  it('checks its input', () => {
    expect(() => parseEntry(5 as unknown as string)).toThrow(InputError)
    expect(() => parseEntry('1', { locale: 'de-DE' })).toThrow(/opts\.locale/)
    expect(() => parseEntry('1', { now: 'yesterday' })).toThrow(/opts\.now: invalid moment/)
    expect(() => parseEntry('1', { cell: { numberFormat: 50 } })).toThrow(/opts\.cell\.numberFormat/)
    expect(() => parseEntry('1', { cell: { style: -1 } })).toThrow(/opts\.cell\.style/)
    expect(() => parseEntry('1', { width: 3 } as never)).toThrow(/unknown/)
  })

  it("puts typed text into a session's cell, read from the session with its style and meta", () => {
    const book = { locale: 'ru-RU', styles: [{ font: { bold: true } }], sheets: [{ name: 'S', cells: { A1: { value: 1, numberFormat: '0.00', style: 0, meta: { id: 7 } }, A2: { formula: 'A1*2' } } }] }
    const session = createSession(book, { clock: { now, seed: 1 } })
    const r = parseEntry('1 234,5', { locale: 'ru-RU', cell: session.getCells('S', 'A1').A1 })
    expect(r).toEqual({ ok: true, input: { value: 1234.5, numberFormat: '0.00', style: 0, meta: { id: 7 } } })
    if (r.ok) expect(session.setCells([{ sheet: 'S', cell: 'A1', input: r.input }]).changes).toEqual([
      { sheet: 'S', cell: 'A1', value: 1234.5, text: '1234,50' },
      { sheet: 'S', cell: 'A2', value: 2469, text: '2469' },
    ])
    session.close()
  })
})
