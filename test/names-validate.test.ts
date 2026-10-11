/**
 * `validateName` in a book's locale and `validateNames`, which gives a book's name issues without the book: the
 * same issues as `ComputeResult.nameIssues`, each telling a name left out of the book from one kept with an error.
 */
import { describe, expect, it } from 'vitest'
import { compute, engine, InputError, openSession, validateName, validateNames, type BookInput, type NameInput } from '../src/index.js'
import './locales.js'

const NAMES: NameInput[] = [
  { name: 'Ставка', ref: '0.2' },
  { name: 'ИСТИНА', ref: '1' },
  { name: 'ложь', ref: '2' },
  { name: 'R', ref: '3' },
  { name: 'C', ref: '4' },
  { name: 'Q1', ref: '5' },
  { name: 'nosheet', ref: 'A1:B2' },
  { name: 'bad', ref: 'SUM(' },
  { name: 'local', ref: "'Data'!A1", scope: 'Nowhere' },
  { name: 'ставка', ref: '0.3' },
  { name: 'gone', ref: "'Other'!A1" },
  { name: 'notable', ref: 'Missing[Qty]' },
  { name: 'udf', ref: 'NOSUCH(1)' },
  { name: 'qty', ref: 'Sales[Qty]' },
]

const book = (locale?: string): BookInput => ({
  ...(locale ? { locale } : {}),
  sheets: [{ name: 'Data', cells: { A1: { value: 'Qty' }, A2: { value: 1 } }, tables: [{ name: 'Sales', range: 'A1:A2', columns: [{ name: 'Qty' }] }] }],
  names: NAMES,
})

const ctx = (locale?: string) => ({ sheets: ['Data'], tables: [{ name: 'Sales', columns: ['Qty'] }], ...(locale ? { locale } : {}) })

/** The names an issue list leaves out, and those it keeps with an error. */
const parts = (issues: { name: string; dropped: boolean }[]) => ({ dropped: issues.filter((x) => x.dropped).map((x) => x.name), kept: issues.filter((x) => !x.dropped).map((x) => x.name) })

describe('validateName in a book\'s locale', () => {
  it('refuses ИСТИНА and ЛОЖЬ in any case in a ru-RU book, takes them in an en-US one', () => {
    for (const name of ['ИСТИНА', 'ложь', 'Ложь']) {
      expect(validateName(name, { locale: 'ru-RU' }).ok, name).toBe(false)
      expect(validateName(name).ok, name).toBe(true)
      expect(validateName(name, { locale: 'en-US' }).ok, name).toBe(true)
    }
    expect(validateName('Ставка', { locale: 'ru-RU' })).toEqual({ ok: true })
  })

  it('gives every locale Excel\'s rules: TRUE, FALSE, R, C and addresses are no names', () => {
    for (const locale of ['en-US', 'ru-RU'])
      for (const name of ['TRUE', 'false', 'R', 'c', 'R2C3', 'Q1', '1x']) expect(validateName(name, { locale }).ok, `${locale} ${name}`).toBe(false)
  })

  it('refuses with a code by the rule a name breaks, and a message in English', () => {
    const code = (name: string, locale?: string) => {
      const r = validateName(name, locale ? { locale } : undefined)
      return r.ok ? 'ok' : r.code
    }
    expect(['1x', 'Став ка', '', 'a'.repeat(256), '?x'].map((n) => code(n))).toEqual(['chars', 'chars', 'chars', 'chars', 'chars'])
    expect(['Q1', 'R', 'c', 'R2C3', 'XFD1'].map((n) => code(n))).toEqual(['reference', 'reference', 'reference', 'reference', 'reference'])
    expect([code('TRUE'), code('false', 'ru-RU'), code('ИСТИНА', 'ru-RU'), code('ИСТИНА')]).toEqual(['boolean', 'boolean', 'boolean', 'ok'])
    expect(validateName('Q1')).toEqual({ ok: false, code: 'reference', message: 'a name must not look like a cell address or an R1C1 reference' })
    expect(validateName('ложь', { locale: 'ru-RU' })).toEqual({ ok: false, code: 'boolean', message: "ложь is a boolean in the book's locale" })
  })

  it('takes a tag of the package\'s locales only', () => {
    expect(() => validateName('x', { locale: 'de-DE' })).toThrow(InputError)
    expect(() => validateName('x', { lang: 'ru-RU' } as never)).toThrow(InputError)
  })

  it('is the engine object\'s too', () => expect(engine.validateName('ложь', { locale: 'ru-RU' }).ok).toBe(false))
})

describe('validateNames', () => {
  it('tells the names left out from those kept with an error', () => {
    expect(parts(validateNames(NAMES, ctx('ru-RU')))).toEqual({
      dropped: ['ИСТИНА', 'ложь', 'R', 'C', 'Q1', 'nosheet', 'bad', 'local', 'ставка'],
      kept: ['gone', 'notable', 'udf'],
    })
    expect(parts(validateNames(NAMES, ctx()))).toEqual({ dropped: ['R', 'C', 'Q1', 'nosheet', 'bad', 'local', 'ставка'], kept: ['gone', 'notable', 'udf'] })
  })

  it('gives the issues a book of these sheets, tables and locale gives, in compute and in a session', () => {
    for (const locale of [undefined, 'en-US', 'ru-RU']) {
      const issues = validateNames(NAMES, ctx(locale))
      expect(compute(book(locale), { display: false }).nameIssues, locale).toEqual(issues)
      const s = openSession(book(locale), { display: false })
      try {
        s.setNames(NAMES.slice(0, 7))
        expect(s.result().nameIssues).toEqual(validateNames(NAMES.slice(0, 7), ctx(locale)))
      } finally { s.close() }
    }
  })

  it('a name kept with an error is a name: a formula finds it, a dropped one is unknown', () => {
    const cells = { B1: { formula: 'IF(FALSE,gone)' }, B2: { formula: 'IF(FALSE,ИСТИНА)' } }
    const r = compute({ locale: 'ru-RU', sheets: [{ name: 'Data', cells }], names: [{ name: 'gone', ref: "'Other'!A1" }, { name: 'ИСТИНА', ref: '1' }] }, { display: false })
    expect(r.sheets[0].cells.B1.value).toBe(false)
    expect(r.sheets[0].cells.B2.value).toEqual({ error: 'NAME' })
  })

  it('needs the book\'s sheets and a locale of the package', () => {
    expect(() => validateNames(NAMES, {} as never)).toThrow(InputError)
    expect(() => validateNames(NAMES, { sheets: ['Data'], locale: 'tr-TR' })).toThrow(InputError)
    expect(() => validateNames([{ name: 'x' }] as never, { sheets: [] })).toThrow(InputError)
    expect(validateNames([], { sheets: [] })).toEqual([])
  })

  it('is the engine object\'s too', () => expect(engine.validateNames([{ name: 'R', ref: '1' }], { sheets: ['S'] })).toEqual([{ name: 'R', reason: expect.any(String), dropped: true }]))
})
