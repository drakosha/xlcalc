/**
 * Locales as modules: en-US is built in, ru-RU is the module `xlcalc/locales/ru-RU`, which `registerLocale` loads. Not
 * loaded, a ru-RU book is an `InputError` that names the module, in every caller that takes a book, and so is the tag
 * wherever a locale is taken; a locale object works without loading where no book is needed; an en-US book shows what
 * it showed (ru-RU's month names under `[$-419]`, the widths of Cyrillic letters); a file's ru-RU locale is read as
 * data. Loaded, a ru-RU book computes as before in both execution modes. The package's lists of its locales agree.
 * The file loads nothing before its part on `registerLocale`: it does not import `locales.ts`.
 */
import { readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { builtinFormatCode, compute, createInlineCalc, displayValue, entryWidth, enUS, fitColumns, InputError, LOCALES, openSession, parseEntry, registerLocale, restructure, toInvariant, toLocal, validateLocalFormula, validateName, validateNames, type BookInput, type Locale } from '../src/index.js'
import { PACKAGE_LOCALES } from '../src/locale/registry.js'
import { ruRU } from '../src/locales/ru-RU.js'
import { LOCALE_MODULES } from '../src/worker/locales.js'
import { writeXlsx } from '../src/xlsx.js'
import { read } from './xlsx-fixtures.js'

const clock = { now: '2026-10-05T12:00:00Z', seed: 1 }
const HINT = "the locale ru-RU is not loaded: import { ruRU } from 'xlcalc/locales/ru-RU' and call registerLocale(ruRU) (loaded: en-US)"
const cells = { A1: { formula: 'VALUE("1,5")+VALUE("1 234,5")' }, A2: { formula: '"a"&1.5&TRUE' }, A3: { value: 46096, format: 14 }, B1: { value: 'Итого' } }
const ruBook: BookInput = { locale: 'ru-RU', sheets: [{ name: 'S', cells, columns: { A: { width: 12 } } }] }
const enBook: BookInput = { sheets: [{ name: 'S', cells }] }

/** The error of a call, which must be an InputError. */
function refusal(f: () => unknown): InputError {
  try {
    f()
  } catch (e) {
    expect(e).toBeInstanceOf(InputError)
    return e as InputError
  }
  throw new Error('expected an InputError')
}

describe('without ru-RU loaded', () => {
  it('LOCALES holds en-US alone, while the package has en-US and ru-RU', () => {
    expect([...LOCALES.keys()]).toEqual(['en-US'])
    expect(LOCALES.get('en-US')).toBe(enUS)
    expect(PACKAGE_LOCALES).toEqual(['en-US', 'ru-RU'])
  })

  it('a ru-RU book is an InputError naming the module to import, in every caller that takes a book', async () => {
    const result = compute(enBook, { clock })
    const calls: [string, () => unknown][] = [
      ['compute', () => compute(ruBook, { clock })],
      ['openSession', () => openSession(ruBook, { clock })],
      ['restructure', () => restructure(ruBook, { op: 'insertRows', sheet: 'S', row: 1 })],
      ['fitColumns', () => fitColumns(ruBook, result, 'S', 'A:B')],
    ]
    for (const [name, call] of calls) expect(refusal(call).message, name).toBe(`book.locale: ${HINT}`)
    const calc = createInlineCalc()
    await expect(calc.compute(ruBook, { clock })).rejects.toThrow(`book.locale: ${HINT}`)
    await expect(calc.openSession(ruBook, { clock })).rejects.toBeInstanceOf(InputError)
    await calc.close()
    await expect(writeXlsx({ book: ruBook })).rejects.toThrow(`book.locale: ${HINT}`)
  })

  it('the tag ru-RU wherever a locale is taken is the same InputError, at its path', () => {
    const calls: [string, () => unknown][] = [
      ['opts.locale', () => parseEntry('1', { locale: 'ru-RU' })],
      ['opts.locale', () => displayValue(1, { locale: 'ru-RU' })],
      ['opts.locale', () => entryWidth(1, { locale: 'ru-RU' })],
      ['opts.locale', () => validateName('x', { locale: 'ru-RU' })],
      ['ctx.locale', () => validateNames([], { sheets: ['S'], locale: 'ru-RU' })],
      ['locale', () => toLocal('1', 'ru-RU')],
      ['locale', () => toInvariant('1', 'ru-RU')],
      ['locale', () => validateLocalFormula('1', 'ru-RU')],
      ['locale', () => builtinFormatCode(14, 'ru-RU')],
    ]
    for (const [path, call] of calls) expect(refusal(call)).toMatchObject({ path, message: `${path}: ${HINT}` })
  })

  it('a locale object works without loading where no book is needed', () => {
    expect(toLocal('IF(B1>0.5,SUM({1,2;3,4}),#N/A)', ruRU)).toEqual({ ok: true, formula: 'ЕСЛИ(B1>0,5;СУММ({1;2:3;4});#Н/Д)' })
    expect(toInvariant('=СУММ(A1:A3;1,5)', ruRU)).toEqual({ ok: true, formula: 'SUM(A1:A3,1.5)' })
    expect(validateLocalFormula('=СУММ(A1;1.5)', ruRU)).toEqual({ ok: false, code: 'syntax', message: 'Здесь не может стоять «.»', at: 10 })
    expect(parseEntry('1 234,5', { locale: ruRU })).toEqual({ ok: true, input: { value: 1234.5, format: 4 } })
    expect(parseEntry('=ОКРУГЛ(A1;2;3)', { locale: ruRU })).toMatchObject({ ok: false, message: 'Функция ОКРУГЛ принимает 2 аргумента, а указано 3' })
    expect(displayValue(0.4567, { format: '0.0%', locale: ruRU })).toEqual({ text: '45,7%' })
    expect(displayValue(46096, { format: 14, locale: ruRU, width: 12 })).toEqual({ text: '15.03.2026' })
    expect(entryWidth(1234567890, { locale: ruRU })).toBe(11.1640625)
    expect(builtinFormatCode(14, ruRU)).toBe('dd.mm.yyyy')
    expect([validateName('ложь', { locale: ruRU }).ok, validateName('ложь').ok]).toEqual([false, true])
    expect(validateNames([{ name: 'ИСТИНА', ref: '1' }], { sheets: ['S'], locale: ruRU })).toMatchObject([{ name: 'ИСТИНА', dropped: true }])
    expect([...LOCALES.keys()]).toEqual(['en-US'])
  })

  it("an en-US book shows what it showed: ru-RU's month names under [$-419], the widths of Cyrillic letters", () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: { A1: { formula: 'TEXT(45688,"[$-419]mmmm")' }, A2: { value: 45688, format: '[$-419]d mmm yyyy' }, B1: { value: 'Итого' }, C1: { value: 'Ж' } }, columns: { A: { width: 16 } } }] }
    const r = compute(book, { clock })
    expect([r.sheets[0].cells.A1.value, r.sheets[0].cells.A2.text]).toEqual(['Январь', '31 янв 2025'])
    // the widths Excel for Mac gives these letters, as in a ru-RU book (test/excel/display/widths.json)
    expect(fitColumns(book, r, 'S', 'B:C')).toEqual({ B: 5.6640625, C: 2.6640625 })
  })

  it("a file's ru-RU locale is read as data, and computing the book names the module", async () => {
    const custom = {
      '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>`,
      'docProps/custom.xml': '<Properties xmlns:vt="v"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="xlcalc.locale"><vt:lpwstr>ru-RU</vt:lpwstr></property></Properties>',
    }
    const r = await read({ sheets: [{ name: 'S', body: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' }], parts: custom })
    expect([r.book.locale, r.issues]).toEqual(['ru-RU', []])
    expect(() => compute(r.book, { clock })).toThrow(`book.locale: ${HINT}`)
  })
})

describe('registerLocale', () => {
  it('loads ru-RU: LOCALES has it, and its tag is taken wherever a locale is', () => {
    registerLocale(ruRU)
    expect([...LOCALES.keys()]).toEqual(['en-US', 'ru-RU'])
    expect(LOCALES.get('ru-RU')).toBe(ruRU)
    expect(toLocal('SUM(A1,1.5)', 'ru-RU')).toEqual({ ok: true, formula: 'СУММ(A1;1,5)' })
    expect(parseEntry('1 234,5', { locale: 'ru-RU' })).toEqual({ ok: true, input: { value: 1234.5, format: 4 } })
  })

  describe.each(['closures', 'eval'] as const)('a ru-RU book computes as before, execution %s', (execution) => {
    it('in compute, a session and the same thread alike', async () => {
      const r = compute(ruBook, { clock, execution })
      const v = r.sheets[0].cells
      expect([v.A1.value, v.A2.value, v.A3.text]).toEqual([1236, 'a1,5ИСТИНА', '15.03.2026'])
      const session = openSession(ruBook, { clock, execution })
      expect(session.result().hash).toBe(r.hash)
      expect(session.explain('S', 'A1').text).toBe("'S'!A1 = 1236  ← ЗНАЧЕН(\"1,5\")+ЗНАЧЕН(\"1 234,5\")")
      expect(session.restructure({ op: 'deleteSheet', sheet: 'S' })).toMatchObject({ ok: false, message: 'В книге должен остаться хотя бы один лист' })
      const calc = createInlineCalc()
      expect((await calc.compute(ruBook, { clock, execution })).hash).toBe(r.hash)
      const inline = await calc.openSession(ruBook, { clock, execution })
      expect([(await inline.result()).hash, (await inline.explain('S', 'A1')).text]).toEqual([r.hash, session.explain('S', 'A1').text])
      await calc.close()
    })
  })

  it('loading the same object again does nothing; another object of a loaded tag, en-US, or a locale a book cannot compute in is an InputError', () => {
    expect(() => registerLocale(ruRU)).not.toThrow()
    const bad = (locale: unknown) => refusal(() => registerLocale(locale as Locale)).message
    expect(bad({ ...ruRU })).toBe('locale.tag: the locale ru-RU is loaded already, as another object')
    expect(bad({ ...enUS })).toBe('locale.tag: the locale en-US is loaded already: it is built in')
    expect(bad({ ...ruRU, tag: 'xx-XX', regional: undefined })).toMatch(/^locale\.regional: expected the rules a book computes by/)
    expect(bad({ ...ruRU, tag: 'xx-XX', regional: { ...ruRU.regional, columnWord: undefined } })).toBe('locale.regional.columnWord: expected a value')
    expect(bad({ ...ruRU, tag: 'xx-XX', regional: { ...ruRU.regional, extra: 1 } })).toMatch(/^locale\.regional: unknown field "extra"/)
    expect(bad({ ...ruRU, tag: 'xx-XX', messages: { lastSheet: 'В книге…' } })).toBe('locale.messages.lastSheet: expected a function')
    expect(bad({ ...ruRU, tag: 'xx-XX', messages: { lastSheets: () => '' } })).toMatch(/^locale\.messages: unknown field "lastSheets"/)
    expect(bad({ ...ruRU, tag: 'xx-XX', listSeparator: 'x' })).toBe('locale.listSeparator: expected one character that has no other meaning in a formula, got "x"')
    expect(bad('ru-RU')).toBe('locale: expected a locale\'s tag or a locale object such as ruRU, got the string "ru-RU"')
    expect([...LOCALES.keys()]).toEqual(['en-US', 'ru-RU'])
  })

  it("a locale of the package's shape under another tag loads too, and its books compute as ru-RU's", () => {
    registerLocale({ ...ruRU, tag: 'xx-XX' })
    const own = compute({ ...ruBook, locale: 'xx-XX' }, { clock }).sheets[0].cells
    expect(own).toEqual(compute(ruBook, { clock }).sheets[0].cells)
  })
})

describe("the package's lists of its locales", () => {
  it('agree: the files of src/locales, PACKAGE_LOCALES, the modules a thread imports and the subpaths of package.json', async () => {
    const files = readdirSync(new URL('../src/locales/', import.meta.url)).map((f) => f.replace(/\.ts$/, ''))
    const tags = files.filter((t) => t !== 'en-US').sort()
    expect([...files].sort()).toEqual([...PACKAGE_LOCALES].sort())
    expect(Object.keys(LOCALE_MODULES).sort()).toEqual(tags)
    const pkg = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8')) as { exports: Record<string, unknown> }
    expect(Object.keys(pkg.exports).filter((k) => k.startsWith('./locales/'))).toEqual(tags.map((t) => `./locales/${t}`))
    // each module exports its locale alone, by the tag without the hyphen, the name the hint of a locale not loaded gives
    for (const tag of tags) {
      const source = readFileSync(new URL(`../src/locales/${tag}.ts`, import.meta.url), 'utf8')
      expect(source.match(/^export .*$/gm), tag).toEqual([`export const ${tag.replace('-', '')}: Locale = {`])
      expect((await LOCALE_MODULES[tag]()).tag).toBe(tag)
    }
  })
})
