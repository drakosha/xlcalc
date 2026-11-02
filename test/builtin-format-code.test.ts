/**
 * `builtinFormatCode`: the code a built-in number format shows by in a locale, the one the engine shows the number by
 * (the same text for every value), undefined for a number that is no built-in format's; its input checks.
 */
import { describe, expect, it } from 'vitest'
import { builtinFormatCode, displayValue, InputError, LOCALES, numberFormatCategory } from '../src/index.js'
import './locales.js'

const VALUES = [0, 1234.567, -1234.5, 0.25, 46096.75, -0.5, 'text']

describe('builtinFormatCode', () => {
  it("gives the code Excel shows a built-in format by in the locale: 14 is the locale's short date", () => {
    expect([builtinFormatCode(0), builtinFormatCode(14), builtinFormatCode(14, 'en-US'), builtinFormatCode(14, 'ru-RU')]).toEqual(['General', 'm/d/yyyy', 'm/d/yyyy', 'dd/mm/yyyy'])
    expect([builtinFormatCode(7, 'ru-RU'), builtinFormatCode(22, 'ru-RU'), builtinFormatCode(49, 'ru-RU')]).toEqual(['#,##0.00 "₽";-#,##0.00 "₽"', 'dd/mm/yyyy h:mm', '@'])
    expect(builtinFormatCode(9, 'ru-RU')).toBe('0%')
  })

  it("gives en-US's money as Excel keeps it in a file, the dollar quoted: the code is the built-in format's category and Excel takes it back as the built-in", () => {
    // written bare, `$` is a literal in a file's code: Excel names such a code Custom and saves it as a format of its own (test/excel/win/ribbon/ribbon-builtin-codes.json)
    expect([5, 6, 7, 8, 42, 44].map((id) => builtinFormatCode(id))).toEqual([
      '"$"#,##0_);("$"#,##0)', '"$"#,##0_);[Red]("$"#,##0)', '"$"#,##0.00_);("$"#,##0.00)', '"$"#,##0.00_);[Red]("$"#,##0.00)',
      '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)', '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
    ])
  })

  it("gives ru-RU's short date and fractions as Excel keeps them in a file: the date's separator `/`, the fraction's space quoted", () => {
    // with dots or a bare space they are a book's own formats, (все форматы) (test/excel/win/ribbon/ru-RU/ribbon-builtin-ru.json)
    expect([12, 13, 14, 22, 27, 36].map((id) => builtinFormatCode(id, 'ru-RU'))).toEqual(['#" "?/?', '#" "??/??', 'dd/mm/yyyy', 'dd/mm/yyyy h:mm', 'dd/mm/yyyy', 'dd/mm/yyyy'])
  })

  it("a built-in format's code has the category of its number in every locale", () => {
    for (const locale of LOCALES.keys())
      for (let id = 0; id < 50; id++) expect(numberFormatCategory(builtinFormatCode(id, locale)!, locale), `${locale} ${id}`).toEqual(numberFormatCategory(id, locale))
  })

  it('undefined for a number that is no built-in numberFormat: a built-in format is one it has a code for', () => {
    for (const id of [-1, 50, 164, 1.5, Number.NaN, Infinity]) expect(builtinFormatCode(id), String(id)).toBeUndefined()
    expect(Array.from({ length: 60 }, (_, id) => builtinFormatCode(id) !== undefined).lastIndexOf(true)).toBe(49)
  })

  it('a value shows by the code as by the number, in every locale and for every built-in format', () => {
    const off: string[] = []
    for (const locale of LOCALES.keys())
      for (let id = 0; id < 50; id++)
        for (const v of VALUES) {
          const [byId, byCode] = [displayValue(v, { numberFormat: id, locale }), displayValue(v, { numberFormat: builtinFormatCode(id, locale), locale })]
          if (JSON.stringify(byId) !== JSON.stringify(byCode)) off.push(`${locale} ${id} ${v}: ${JSON.stringify(byId)} by the number, ${JSON.stringify(byCode)} by the code`)
        }
    expect(off).toEqual([])
  })

  it('checks its input', () => {
    expect(() => builtinFormatCode('14' as never)).toThrow(new InputError('id', 'expected the number of a built-in format, 0 to 49, got the string "14"'))
    expect(() => builtinFormatCode(14, 'de-DE')).toThrow(new InputError('locale', `unknown locale "de-DE" (loaded: ${[...LOCALES.keys()].join(', ')}; the package's: en-US, ru-RU)`))
  })
})
