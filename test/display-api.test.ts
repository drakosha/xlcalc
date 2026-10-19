/**
 * The display of values in results and `displayValue`: on by default and off on request, values unchanged and
 * the hash without it, formats of rows and columns, merged areas, the book's locale, the colours the screen shows
 * where Excel's AppleScript reads none, input checks.
 */
import { describe, expect, it } from 'vitest'
import type { BookInput } from '../src/contract.js'
import { compute, displayValue, entryWidth, fitColumns, InputError, LOCALES } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import './locales.js'

const clock = { now: '2026-10-07T12:00:00Z', seed: 1 }

describe('results show what the cells show', () => {
  const book: BookInput = {
    sheets: [
      {
        name: 'S',
        cells: { A1: { value: 2, numberFormat: '0.00' }, A2: { value: 1234567.891 }, A3: { formula: 'A1*-3', numberFormat: '[Blue]0;[Red]-0' }, B1: { value: 'text' }, B2: { formula: 'SEQUENCE(2)' }, C5: { value: 7 } },
        columns: { A: { width: 6 }, B: { numberFormat: '0.0' } },
        rows: { 5: { numberFormat: '0%' } },
      },
    ],
  }

  it('by default: text and colour beside the value, which stays a number', () => {
    const cells = compute(book, { clock }).sheets[0].cells
    expect(cells.A1).toEqual({ value: 2, text: '2.00', numberFormat: '0.00' })
    expect(cells.A2).toEqual({ value: 1234567.891, text: '1E+06' })
    expect(cells.A3).toEqual({ value: -6, text: '-6', color: '#FF0000', formula: 'A1*-3', numberFormat: '[Blue]0;[Red]-0' })
    expect(cells.B1.text).toBe('text')
  })

  it("a cell without a format shows its row's, else its column's; spilled cells too", () => {
    const cells = compute(book, { clock }).sheets[0].cells
    expect([cells.B2.text, cells.B3.text, cells.C5.text]).toEqual(['1.0', '2.0', '700%'])
  })

  it('display: false leaves the text out; the values and the hash do not depend on it', () => {
    const on = compute(book, { clock })
    const off = compute(book, { clock, display: false })
    expect(Object.values(off.sheets[0].cells).some((c) => 'text' in c || 'color' in c)).toBe(false)
    expect(off.hash).toBe(on.hash)
    expect(off.sheets[0].cells.A3).toEqual({ value: -6, formula: 'A1*-3', numberFormat: '[Blue]0;[Red]-0' })
  })

  it("a merged area shows its top-left value in its columns' width", () => {
    const sheet = (merges?: string[]): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 123456789 } }, columns: { A: { width: 4 }, B: { width: 4 } }, ...(merges ? { merges } : {}) }] })
    expect(compute(sheet(), { clock }).sheets[0].cells.A1.text).toBe('###')
    expect(compute(sheet(['A1:B1']), { clock }).sheets[0].cells.A1.text).toBe('1.2E+08')
  })

  it("in the book's locale", () => {
    const ru = compute({ ...book, locale: 'ru-RU' }, { clock }).sheets[0].cells
    expect([ru.A1.text, ru.B2.text]).toEqual(['2,00', '1,0'])
    const dates = compute({ locale: 'ru-RU', sheets: [{ name: 'S', cells: { A1: { value: { date: '2025-01-31' }, numberFormat: 'dd/mm/yyyy' }, A2: { value: true } } }] }, { clock }).sheets[0].cells
    // the date is wider than the default column
    expect([dates.A1.text, dates.A2.text]).toEqual(['#########', 'ИСТИНА'])
    expect(displayValue(45688, { numberFormat: 'dd/mm/yyyy', locale: 'ru-RU', width: 12 }).text).toBe('31.01.2025')
  })
})

describe('displayValue', () => {
  it('shows a value by a format, a locale and a width', () => {
    expect(displayValue(1234.567)).toEqual({ text: '1234.567' })
    expect(displayValue(1234.567, { numberFormat: '#,##0.00', locale: 'ru-RU' })).toEqual({ text: '1 234,57' })
    expect(displayValue(-5, { numberFormat: '[Blue]0;[Red]-0' })).toEqual({ text: '-5', color: '#FF0000' })
    expect(displayValue(123456789, { width: 3 })).toEqual({ text: '##' })
    expect(displayValue(45688.75, { numberFormat: 'h:mm AM/PM' })).toEqual({ text: '6:00 PM' })
    expect(displayValue({ error: 'DIV0' }, { locale: 'ru-RU' })).toEqual({ text: '#ДЕЛ/0!' })
    expect(displayValue(null, { numberFormat: '0.00' })).toEqual({ text: '' })
    expect(displayValue('abc', { numberFormat: '"<"@">"' })).toEqual({ text: '<abc>' })
  })

  it('without a width the width is not limited: General keeps its digits, no number turns into #', () => {
    expect(displayValue(46096, { numberFormat: 'dd.mm.yyyy' })).toEqual({ text: '15.03.2026' })
    expect([displayValue(123456789).text, displayValue(1 / 3).text, displayValue(1234567890123456).text]).toEqual(['123456789', '0.333333333', '1.23457E+15'])
  })

  it("a column of the default width, given as xlsx stores it, narrows as before", () => {
    const width = 8.83203125
    expect([displayValue(12345678, { width }).text, displayValue(123456789, { width }).text]).toEqual(['12345678', '1.23E+08'])
    expect(displayValue(46096, { numberFormat: 'dd.mm.yyyy', width }).text).toBe('#########')
  })

  it('entryWidth, fitColumns and the results keep their own widths: the default one without a width', () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 46096, numberFormat: 'dd.mm.yyyy' }, B1: { value: 123456789 } } }] }
    const r = compute(book, { clock })
    expect([r.sheets[0].cells.A1.text, r.sheets[0].cells.B1.text]).toEqual(['#########', '1.23E+08'])
    expect(fitColumns(book, r, 'S', 'A:B')).toEqual({ A: 10.1640625, B: 10.1640625 })
    expect(entryWidth(1234567890)).toBe(11.1640625)
    expect(entryWidth(1234567890, { width: 9 })).toBeUndefined()
  })

  it('checks its input', () => {
    const fails = (f: () => unknown, message: string) => expect(f).toThrow(new InputError(message.split(':')[0], message.slice(message.indexOf(':') + 2)))
    fails(() => displayValue(1, { locale: 'de-DE' }), `opts.locale: unknown locale "de-DE" (loaded: ${[...LOCALES.keys()].join(', ')}; the package's: en-US, ru-RU)`)
    fails(() => displayValue(1, { width: -1 }), 'opts.width: expected a number from 0 to 4177920, got number -1')
    fails(() => displayValue(1, { size: 3 } as never), 'opts: unknown field "size" (expected numberFormat, locale, width, dateSystem)')
    fails(() => displayValue(Number.NaN), 'value: expected a number, text, a boolean, null or { error: code }, got number NaN')
    fails(() => displayValue({ error: 'OOPS' } as never), 'value: expected a number, text, a boolean, null or { error: code }, got an object')
    fails(() => compute({ sheets: [] }, { display: 'yes' } as never), 'opts.display: expected a boolean, got the string "yes"')
  })
})

describe("the system's long date and time by their culture tags", () => {
  const LONG = '[$-x-sysdate]dddd, mmmm dd, yyyy'
  const TIME = '[$-x-systime]h:mm:ss AM/PM'

  it('show as [$-F800] and [$-F400] do, whatever the code after the tag; a currency shows its symbol, not the culture', () => {
    const shown = (v: number, numberFormat: string, locale?: string) => displayValue(v, { numberFormat, locale }).text
    expect([LONG, '[$-x-sysdate]yyyy-mm-dd', TIME, '[$-X-SYSTIME]yyyy'].map((f) => shown(45688.5, f))).toEqual(['Friday, January 31, 2025', 'Friday, January 31, 2025', '12:00:00 PM', '12:00:00 PM'])
    expect([LONG, TIME].map((f) => shown(45001.25, f, 'ru-RU'))).toEqual(['четверг, 16 марта 2023 г.', '6:00:00'])
    expect(['[$€-x-euro2] #,##0.00', '#,##0.00 [$₽-ru-RU]', '[$$-en-US]#,##0.00'].map((f) => shown(1.5, f))).toEqual(['€ 1.50', '1.50 ₽', '$1.50'])
  })

  it('in a file read and computed', async () => {
    const { data } = await writeXlsx({ book: { sheets: [{ name: 'S', cells: { A1: { value: 45688.5, numberFormat: LONG }, A2: { value: 45688.5, numberFormat: TIME } }, columns: { A: { width: 40 } } }] } })
    const cells = compute((await readXlsx(data)).book, { clock }).sheets[0].cells
    expect([cells.A1.text, cells.A2.text]).toEqual(['Friday, January 31, 2025', '12:00:00 PM'])
  })
})

describe("colours the screen shows where Excel's AppleScript reads none", () => {
  // read from a screenshot of Excel 16.113 for Mac (docs/formats.md): the section's colour on a zero and on text
  it.each([
    [0, '[Red]0', '#FF0000'],
    [0, '[Red]0;[Blue]0', '#FF0000'],
    [0, '[Red][=0]0;[Blue]0', '#FF0000'],
    ['abc', '[Red]@', '#FF0000'],
    ['abc', '0;0;0;[Red]@', '#FF0000'],
    ['x', '[Red]0;[Blue]0;[Green]0;[Magenta]@', '#FF00FF'],
    [-5, '[Blue]0;[Red]0', '#FF0000'],
  ] as const)('%j by %s', (value, format, color) => expect(displayValue(value, { numberFormat: format }).color).toBe(color))
})
