/**
 * Every function that takes a locale takes the tag of a loaded one or the locale object alike: the same answer for
 * `'ru-RU'` and `ruRU`, for `'en-US'` and `enUS` (and for none, en-US). A tag is compared as written: another tag is an
 * `InputError` that lists the loaded ones, and names the tag a different case meant. A locale object of formula text
 * alone, without rules of computation, translates in its language and reads and shows values as en-US does.
 */
import { describe, expect, it } from 'vitest'
import { builtinFormatCode, displayValue, entryWidth, enUS, InputError, parseEntry, toInvariant, toLocal, validateLocalFormula, validateName, validateNames, type Locale } from '../src/index.js'
import { ruRU } from './locales.js'

const now = '2026-10-10T12:00:00Z'

/** Each function with a locale, called with one. */
const CALLS: [string, (locale: string | Locale) => unknown][] = [
  ['toLocal', (l) => [toLocal('IF(B1>0.5,SUM({1,2;3,4}),#N/A)', l), toLocal('SUM(Sales[[#This Row],[Qty]])', l, { table: 'Sales' })]],
  ['toInvariant', (l) => [toInvariant('=СУММ(A1:A3;1,5)', l), toInvariant('=SUM(A1:A3,1.5)', l)]],
  ['validateLocalFormula', (l) => [validateLocalFormula('=ОКРУГЛ(Данные!B2*1,2;0)', l, { sheets: ['Данные'] }), validateLocalFormula('=СУММ(A1;1.5)', l)]],
  ['parseEntry', (l) => ['1 234,5', '15.03.2026', '3/15/2026', 'ИСТИНА', '=ОКРУГЛ(A1;2;3)'].map((t) => parseEntry(t, { locale: l, now }))],
  ['displayValue', (l) => [displayValue(0.4567, { format: '0.0%', locale: l }), displayValue(46096, { format: 14, locale: l, width: 12 })]],
  ['entryWidth', (l) => [entryWidth(1234567890, { locale: l }), entryWidth(46096, { format: 14, locale: l })]],
  ['validateName', (l) => [validateName('ложь', { locale: l }), validateName('FALSE', { locale: l })]],
  ['builtinFormatCode', (l) => [builtinFormatCode(14, l), builtinFormatCode(7, l)]],
  ['validateNames', (l) => validateNames([{ name: 'ИСТИНА', ref: '1' }, { name: 'rate', ref: '0.2' }], { sheets: ['S'], locale: l })],
]

describe.each(CALLS)('%s', (_, call) => {
  it('the same answer for the tag and the object', () => {
    expect(call('ru-RU')).toEqual(call(ruRU))
    expect(call('en-US')).toEqual(call(enUS))
    expect(call('ru-RU')).not.toEqual(call('en-US'))
  })

  it('a tag not loaded lists the loaded ones; a tag in another case names the right one', () => {
    const fails = (tag: unknown) => {
      try {
        call(tag as string)
      } catch (e) {
        expect(e).toBeInstanceOf(InputError)
        return (e as InputError).message
      }
      throw new Error('expected an InputError')
    }
    expect(fails('de-DE')).toMatch(/locale: unknown locale "de-DE" \(loaded: en-US, ru-RU; the package's: en-US, ru-RU\)$/)
    expect(fails('ru-ru')).toMatch(/locale: unknown locale "ru-ru" \(loaded: en-US, ru-RU; tags are case-sensitive: ru-RU\)$/)
    expect(fails(42)).toMatch(/locale: expected a locale's tag or a locale object such as ruRU, got number 42$/)
  })
})

describe('a locale of formula text alone', () => {
  const deDE: Locale = { tag: 'de-DE', decimalSeparator: ',', listSeparator: ';', arrayColumnSeparator: '.', arrayRowSeparator: ';', booleans: { TRUE: 'WAHR', FALSE: 'FALSCH' }, errors: { '#N/A': '#NV' }, functions: { SUM: 'SUMME', IF: 'WENN' } }

  it('translates formulas in its language, and reads and shows values as en-US', () => {
    expect(toLocal('IF(TRUE,SUM(1.5,2),#N/A)', deDE)).toEqual({ ok: true, formula: 'WENN(WAHR;SUMME(1,5;2);#NV)' })
    expect(parseEntry('=SUMME(1,5;2)', { locale: deDE })).toEqual({ ok: true, input: { formula: 'SUM(1.5,2)' } })
    expect(parseEntry('1,234.5', { locale: deDE })).toEqual(parseEntry('1,234.5'))
    expect(displayValue(1234.5, { format: '#,##0.00', locale: deDE })).toEqual({ text: '1,234.50' })
    expect(validateName('WAHR', { locale: deDE }).ok).toBe(false)
  })
})
