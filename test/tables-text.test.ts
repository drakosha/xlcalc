/**
 * A formula typed in a cell of a table, without the book: `toInvariant(text, locale, { table })` names the table before a
 * structured reference without one, as `toLocal` with the table drops it, and `validateFormula` and
 * `validateLocalFormula` with `ctx.table` check such a reference as in a cell of the table, the canonical text the one a
 * book stores; in en-US and ru-RU.
 */
import { describe, expect, it } from 'vitest'
import { compute, enUS, InputError, toInvariant, toLocal, validateFormula, validateLocalFormula, type Locale } from '../src/index.js'
import { ruRU } from './locales.js'

const tables = [{ name: 'Sales', columns: ['Qty', 'Unit Price'] }, { name: 'Other', columns: ['Price'] }]

/** Stored formulas of a cell of Sales: its own references without its name once shown, another table's with it. */
const STORED = ['Sales[[#This Row],[Qty]]*Sales[[#This Row],[Unit Price]]', 'SUM(Sales[Qty])/ROWS(Sales[#All])', 'Sales[[#This Row],[Qty]:[Unit Price]]', 'Other[[#This Row],[Price]]+SUM(Sales[[#Totals],[Qty]])']

describe.each([
  ['en-US', enUS],
  ['ru-RU', ruRU],
] as [string, Locale][])('%s', (_, locale) => {
  it('toLocal with the table, then toInvariant with it, reads back to the stored text', () => {
    for (const f of STORED) {
      const shown = toLocal(f, locale, { table: 'Sales' })
      expect(shown.ok, f).toBe(true)
      const back = toInvariant(`=${shown.ok && shown.formula}`, locale, { table: 'Sales' })
      expect(back.ok && validateFormula(back.formula, { tables }), f).toMatchObject({ ok: true, canonical: f })
      // the typed text checked as in the table, without translating it first
      expect(validateLocalFormula(`=${shown.ok && shown.formula}`, locale, { table: 'Sales', tables }), f).toMatchObject({ ok: true, canonical: f })
    }
  })
})

describe('toInvariant with the table of the formula', () => {
  it('names the table before a reference without a name, the rest as toInvariant writes it', () => {
    expect(toInvariant('=[@Qty]*1,5', ruRU, { table: 'Sales' })).toEqual({ ok: true, formula: 'Sales[@Qty]*1.5' })
    expect(toInvariant('=СУММ([[#Итоги];[Qty]])+Other[@Price]', ruRU, { table: 'Sales' })).toEqual({ ok: true, formula: 'SUM(Sales[[#Totals],[Qty]])+Other[@Price]' })
    expect(toInvariant('=SUM([Qty]) + [@[Unit Price]]', enUS, { table: 'Sales' })).toEqual({ ok: true, formula: 'SUM(Sales[Qty]) + Sales[@[Unit Price]]' })
    // without the table, and in text that does not read, the references stay as typed
    expect(toInvariant('=[@Qty]*1,5', ruRU)).toEqual({ ok: true, formula: '[@Qty]*1.5' })
    expect(toInvariant('=[@Qty]*(1,5', ruRU, { table: 'Sales' })).toEqual({ ok: true, formula: 'Sales[@Qty]*(1.5' })
    expect(toInvariant('=[@Qty]+"x', enUS, { table: 'Sales' })).toMatchObject({ ok: false, code: 'syntax', at: 8 })
  })

  it("a wrong option is an InputError", () => {
    expect(() => toInvariant('=[@Qty]', enUS, { table: 1 } as never)).toThrow(new InputError('opts.table', 'expected a string, got number 1'))
  })
})

describe('validateFormula and validateLocalFormula with the table of the formula', () => {
  it('a reference without a table name is to the table, and the canonical text names it', () => {
    expect(validateFormula('[@Qty]*2', { table: 'Sales' })).toEqual({ ok: true, canonical: 'Sales[[#This Row],[Qty]]*2', functions: [], sheets: [], names: [], externals: [], tables: ['Sales'] })
    expect(validateFormula('SUM([qty])', { table: 'sales', tables })).toMatchObject({ ok: true, canonical: 'SUM(Sales[Qty])', tables: ['Sales'] })
    expect(validateLocalFormula('=[@Qty]*1,5', ruRU, { table: 'Sales' })).toMatchObject({ ok: true, canonical: 'Sales[[#This Row],[Qty]]*1.5' })
    // a column the table lacks, a table the book lacks; positions in the typed text
    expect(validateLocalFormula('=1+[@Price]', ruRU, { table: 'Sales', tables })).toMatchObject({ ok: false, code: 'unknown_name', at: 5 })
    expect(validateFormula('[@Qty]', { table: 'Gone', tables })).toMatchObject({ ok: false, code: 'unknown_name', message: 'unknown table Gone' })
    expect(validateLocalFormula('=[@Qty]*1.5', ruRU, { table: 'Sales' })).toMatchObject({ ok: false, code: 'syntax', at: 9 })
  })

  it('without the table as before: refused with the tables, unnamed without them', () => {
    expect(validateFormula('[@Qty]', { tables })).toMatchObject({ ok: false, code: 'unknown_name', at: 0 })
    expect(validateFormula('[@Qty]')).toMatchObject({ ok: true, canonical: '[[#This Row],[Qty]]', tables: [] })
  })

  it("the canonical text is the one a book stores for the table's cell", () => {
    const cells = { A1: { value: 'Qty' }, B1: { value: 'Unit Price' }, A2: { value: 2 }, B2: { value: 3 }, C1: { value: 'Total' }, C2: { formula: '[@Qty]*[@[Unit Price]]+SUM([Qty])' } }
    const r = compute({ sheets: [{ name: 'S', cells, tables: [{ name: 'Sales', range: 'A1:C2', columns: [{ name: 'Qty' }, { name: 'Unit Price' }, { name: 'Total' }] }] }] })
    const v = validateFormula('[@Qty]*[@[Unit Price]]+SUM([Qty])', { table: 'Sales' })
    expect(v.ok && v.canonical).toBe(r.sheets[0].cells.C2.formula)
  })

  it('a wrong context is an InputError', () => {
    expect(() => validateFormula('[@Qty]', { table: ['Sales'] } as never)).toThrow(new InputError('ctx.table', 'expected a string, got an array'))
  })
})
