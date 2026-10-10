/**
 * AutoFit and the width typing gives a column, through the API: `fitColumns` over a computed book, the layout edit
 * `{ columns, fit: true }` in a session (both execution modes, in the same thread and in a thread), `entryWidth`, the
 * column's `bestFit` in xlsx, and the checks of their input. The rules themselves are checked against Excel in
 * `test/fit-excel.test.ts`.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, createInlineCalc, entryWidth, fitColumns, InputError, openSession, parseEntry, type BookInput, type Calc } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { pool } from './fixtures/pool.js'

const clock = { now: '2026-10-10T12:00:00Z', seed: 1 }

const book = (): BookInput => ({
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 'Name' },
        A2: { value: 'Alexander the Great' },
        B1: { value: 46096, format: 14 },
        B2: { value: 1234567.891, format: '#,##0.00' },
        C1: { value: 123456789012 },
        D2: { value: 'merged across' },
        E1: { value: 'Итого' },
      },
      columns: { C: { width: 3, hidden: true }, F: { width: 20 } },
      merges: ['D2:E2'],
      rows: { 2: { hidden: true } },
    },
  ],
})

/** The widths AutoFit gives columns A to F of the book: a hidden row counts, a cell merged across columns does not, F has nothing to show. */
const FITTED = { A: 16, B: 11.6640625, C: 11.83203125, E: 5.6640625 }

describe('fitColumns', () => {
  it('fits the columns of a computed book by its values, formats and locale', () => {
    const b = book()
    expect(fitColumns(b, compute(b, { clock }), 'S', 'A:F')).toEqual(FITTED)
    const ru = { ...book(), locale: 'ru-RU' }
    expect(fitColumns(ru, compute(ru, { clock }), 's', 'B')).toEqual({ B: 11.33203125 })
  })
  it('gives null where AutoFit comes to the default width, and stops at 255', () => {
    const b: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 'bbbbbbbr' }, B1: { value: 'x'.repeat(400) } } }] }
    expect(fitColumns(b, compute(b, { clock }), 'S', 'A:B')).toEqual({ A: null, B: 255 })
  })
  it('refuses an unknown sheet, columns that are none and a result without the sheet', () => {
    const b = book()
    const r = compute(b, { clock })
    expect(() => fitColumns(b, r, 'T', 'A')).toThrow(InputError)
    expect(() => fitColumns(b, r, 'S', 'A1')).toThrow(/columns such as/)
    expect(() => fitColumns(b, { ...r, sheets: [] }, 'S', 'A')).toThrow(/no sheet/)
  })
})

describe.each(['closures', 'eval'] as const)('AutoFit in a session, execution %s', (execution) => {
  it('sets the widths as Excel fitted them, shows a hidden column, returns the cells shown otherwise', () => {
    const s = openSession(book(), { clock, execution })
    expect(s.setLayout([{ sheet: 'S', columns: 'A:F', fit: true }])).toEqual([
      { sheet: 'S', cell: 'B1', value: 46096, text: '3/15/2026' },
      { sheet: 'S', cell: 'C1', value: 123456789012, text: '1.23457E+11' },
      { sheet: 'S', cell: 'B2', value: 1234567.891, text: '1,234,567.89' },
    ])
    expect(s.layout().sheets[0].columns).toEqual({ A: { width: 16, bestFit: true }, B: { width: 11.6640625, bestFit: true }, C: { width: 11.83203125, bestFit: true }, E: { width: 5.6640625, bestFit: true }, F: { width: 20 } })
    expect(compute(s.book(), { clock, execution }).sheets[0].cells.C1.text).toBe('1.23457E+11')
  })
  it('fits by the layout the edits before it leave', () => {
    const s = openSession(book(), { clock, execution })
    s.setLayout([
      { sheet: 'S', columns: 'C', set: { format: '#,##0' } },
      { sheet: 'S', merges: [] },
      { sheet: 'S', columns: 'C:D', fit: true },
    ])
    expect(s.layout().sheets[0].columns).toMatchObject({ C: { width: 14.6640625, bestFit: true, format: '#,##0' }, D: { width: 11.83203125, bestFit: true } })
  })
  it('refuses AutoFit of rows and a fit that is not true, and checks bestFit', () => {
    const s = openSession(book(), { clock, execution })
    expect(() => s.setLayout([{ sheet: 'S', rows: '1', fit: true } as never])).toThrow(InputError)
    expect(() => s.setLayout([{ sheet: 'S', columns: 'A', fit: false } as never])).toThrow(/expected true/)
    expect(() => s.setLayout([{ sheet: 'S', rows: '1', set: { bestFit: true } } as never])).toThrow(InputError)
    expect(() => s.setLayout([{ sheet: 'S', columns: 'A', set: { bestFit: 1 } } as never])).toThrow(InputError)
  })
})

const calcs: Calc[] = []
afterAll(() => Promise.all(calcs.map((c) => c.close())))

describe.each([
  ['in the same thread', async () => createInlineCalc()],
  ['in a thread', async () => pool({ threads: 1 })],
] as const)('AutoFit in an asynchronous session, %s', (_, make) => {
  it('goes through setLayout and the layout gives the widths', async () => {
    const calc = await make()
    calcs.push(calc)
    const s = await calc.openSession(book(), { clock })
    expect((await s.setLayout([{ sheet: 'S', columns: 'A:F', fit: true }])).map((c) => c.cell)).toEqual(['B1', 'C1', 'B2'])
    expect((await s.layout()).sheets[0].columns?.C).toEqual({ width: 11.83203125, bestFit: true })
  })
})

describe('entryWidth', () => {
  const typed = (text: string, opts: { width?: number; bestFit?: boolean } = {}) => {
    const r = parseEntry(text, { locale: 'ru-RU', now: clock.now })
    const v = r.ok ? r.input?.value : undefined
    if (!r.ok || !r.input || v === undefined || (v !== null && typeof v === 'object')) throw new Error(text)
    return entryWidth(v, { locale: 'ru-RU', ...(r.input.format !== undefined ? { format: r.input.format } : {}), ...opts })
  }
  it('widens a column of the default width or one Excel fitted for a number AutoFit gives more room', () => {
    expect(typed('15.03.2026')).toBe(10.1640625)
    expect(typed('1234567890')).toBe(11.1640625)
    expect(typed('15.03.2026', { width: 5, bestFit: true })).toBe(10.1640625)
    expect(entryWidth(12345678)).toBe(9.1640625)
    expect(entryWidth(46096, { format: 14, width: 3, bestFit: true })).toBe(9.5)
  })
  it('keeps a width set by hand, General that drops decimals, text and numbers that fit', () => {
    expect(typed('15.03.2026', { width: 5 })).toBeUndefined()
    expect(typed('3,14159265')).toBeUndefined()
    expect(typed('abc')).toBeUndefined()
    expect(typed('1 234')).toBeUndefined()
    expect(entryWidth(true)).toBeUndefined()
    expect(entryWidth(46096, { format: 14, width: 30, bestFit: true })).toBeUndefined()
  })
  it('keeps a fitted width that equals the default one', () => {
    expect(entryWidth(1234567, { format: 3, locale: 'ru-RU', width: 3, bestFit: true })).toBe(8.83203125)
  })
  it('checks its options', () => {
    expect(() => entryWidth(1, { bestFit: 'yes' } as never)).toThrow(InputError)
    expect(() => entryWidth(1, { width: -1 })).toThrow(InputError)
    expect(() => entryWidth(1, { color: 1 } as never)).toThrow(InputError)
  })
})

describe('bestFit in xlsx', () => {
  it('is written with the width and read back', async () => {
    const b: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 } }, columns: { A: { width: 12.5, bestFit: true }, B: { width: 4 } } }] }
    const { data } = await writeXlsx(b)
    expect((await readXlsx(data)).book.sheets[0].columns).toEqual({ A: { width: 12.5, bestFit: true }, B: { width: 4 } })
  })
})
