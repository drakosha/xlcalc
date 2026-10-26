/**
 * `validateName` in a book's locale and `validateNames`, which gives a book's name issues without the book: the
 * same issues as `ComputeResult.diagnostics` of kind `name`, each telling a name left out of the book from one kept with an error.
 */
import { describe, expect, it } from 'vitest'
import { compute, InputError, createSession, validateName, validateNames, type BookInput, type NameInput } from '../src/index.js'
import './locales.js'
import { nameIssues } from './diagnostics.js'

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
    expect(validateName('ложь', { locale: 'ru-RU' })).toEqual({ ok: false, code: 'boolean', message: "ложь is a boolean in the workbook's locale" })
  })

  it('takes a tag of the package\'s locales only', () => {
    expect(() => validateName('x', { locale: 'de-DE' })).toThrow(InputError)
    expect(() => validateName('x', { lang: 'ru-RU' } as never)).toThrow(InputError)
  })
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
      expect(nameIssues(compute(book(locale), { display: false })), locale).toEqual(issues)
      const s = createSession(book(locale), { display: false })
      try {
        s.setNames(NAMES.slice(0, 7))
        expect(nameIssues(s.getResult())).toEqual(validateNames(NAMES.slice(0, 7), ctx(locale)))
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
})

describe('setNames refuses a name the workbook would drop', () => {
  it('as validateNames finds it, at the name, in setNames and in an operation of applyChanges, the book unchanged', () => {
    for (const locale of [undefined, 'ru-RU']) {
      const s = createSession({ ...(locale ? { locale } : {}), sheets: [{ name: 'Data', cells: { A1: { value: 1 } } }] }, { display: false })
      try {
        const names: NameInput[] = [{ name: 'ok', ref: "'Data'!A1" }, ...['1abc', 'Q1', 'TRUE', ...(locale ? ['ЛОЖЬ'] : [])].map((name) => ({ name, ref: "'Data'!A1" }))]
        for (let i = 1; i < names.length; i++) {
          const given = [names[0], names[i]]
          const [issue] = validateNames(given, { sheets: ['Data'], ...(locale ? { locale } : {}) })
          expect(() => s.setNames(given)).toThrow(new InputError('names[1].name', issue.reason))
          expect(() => s.applyChanges([{ op: 'setNames', names: given }])).toThrow(new InputError('operations[0].names[1].name', issue.reason))
        }
        for (const [given, path] of [
          [[{ name: 'x', ref: "'Data'!A1" }, { name: 'X', ref: '2' }], 'names[1].name'],
          [[{ name: 'x', ref: "'Data'!A1", scope: 'Nowhere' }], 'names[0].name'],
          [[{ name: 'x', ref: 'A1' }], 'names[0].name'],
          [[{ name: 'x', ref: 'SUM(' }], 'names[0].name'],
        ] as [NameInput[], string][])
          expect(() => s.setNames(given), JSON.stringify(given)).toThrow(path)
        expect(s.getBook().names ?? []).toEqual([])
        // a name kept with an error is taken
        s.setNames([{ name: 'udf', ref: 'NOSUCH(1)' }])
        expect(s.getBook().names).toEqual([{ name: 'udf', ref: 'NOSUCH(1)' }])
      } finally { s.close() }
    }
  })

  it('takes back a name the book was opened with as it stands, refuses it changed', () => {
    const had: NameInput[] = [{ name: 'Q1', ref: "'Data'!A1", hidden: true }, { name: 'dup', ref: '1' }, { name: 'DUP', ref: '2' }]
    const s = createSession({ sheets: [{ name: 'Data', cells: {} }], names: had }, { display: false })
    try {
      s.setNames([...had, { name: 'more', ref: '3' }])
      expect(s.getBook().names).toHaveLength(4)
      expect(() => s.setNames([{ ...had[0], ref: "'Data'!A2" }])).toThrow('names[0].name')
      expect(() => s.setNames([had[1], { name: 'Dup', ref: '2' }])).toThrow('names[1].name')
    } finally { s.close() }
  })
})
