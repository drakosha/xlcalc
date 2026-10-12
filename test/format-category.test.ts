/**
 * `numberFormatCategory` and `numberFormatChoices`: the category of a format as the ribbon's Number Format box shows it
 * and the items of its list, in en-US and ru-RU; the code as Excel keeps it (`canonical`); their input checks. Excel's
 * answers are in `test/ui-excel.test.ts` (the ui set `ribbon`, Excel for Mac) and `test/format-category-win.test.ts` (Excel for Windows).
 */
import { describe, expect, it } from 'vitest'
import { canonical } from '../src/numfmt/category.js'
import { InputError, LOCALES, numberFormatCategory, numberFormatChoices, type FormatCategory } from '../src/index.js'
import { ruRU } from '../src/locales/ru-RU.js'
import './locales.js'

const categories = (codes: (string | number)[], locale?: string) => codes.map((c) => numberFormatCategory(c, locale).category)

describe('canonical', () => {
  it('keeps a literal written bare escaped, quoted text and brackets as written, as Excel saves a code', () => {
    expect(['$#,##0_);($#,##0)', '"$"#,##0.00', 'm/d/yyyy h:mm', 'dd.mm.yyyy', '# ?/2', '0.00 "kg"', '[$€-407] #,##0.00', '[h]:mm:ss', 'mm:ss.0'].map(canonical)).toEqual([
      '\\$#,##0_);\\(\\$#,##0\\)', '"$"#,##0.00', 'm/d/yyyy\\ h:mm', 'dd\\.mm\\.yyyy', '#\\ ?/2', '0.00\\ "kg"', '[$€-407]\\ #,##0.00', '[h]:mm:ss', 'mm:ss.0',
    ])
    expect(canonical('[$-x-sysdate]dddd, mmmm dd, yyyy')).toBe('[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy')
  })
})

describe('numberFormatCategory in en-US', () => {
  it("names the category of Format Cells whose formats hold the code: decimals, groups and Excel's four negatives are Number", () => {
    expect(categories(['General', '0', '0.000', '#,##0.0', '0.00;[Red]0.00', '0.00_);(0.00)', '#,##0_);[Red](#,##0)'])).toEqual(['general', ...Array(6).fill('number')])
    expect(categories(['0.00;-0.00', '0.00;[Red]-0.00', '[Red]0.00', '#.##', '00.00', '0.00 "kg"'])).toEqual(Array(6).fill('custom'))
    expect(numberFormatCategory('0.000')).toEqual({ category: 'number', name: 'Number' })
  })

  it('percent, scientific and fraction in the forms Format Cells writes them', () => {
    expect(categories(['0%', '0.0%', '0.000E+00', '0.0E+00', '# ?/?', '# ???/???', '# ?/8', '# ??/100'])).toEqual(['percentage', 'percentage', 'scientific', 'scientific', 'fraction', 'fraction', 'fraction', 'fraction'])
    expect(categories(['#,##0%', '0E+00', '0.0E+0', '?/?', '# ?/3', '#" "?/?'])).toEqual(Array(6).fill('custom'))
  })

  it('money: the quoted dollar or a bracketed symbol of any place is Currency, a dollar written bare a literal', () => {
    expect(categories(['"$"#,##0.00', '[$€-407] #,##0.00', '#,##0.00 [$₽-419]', '[$USD] #,##0.00', '[$$-409]#,##0.00;[Red]-[$$-409]#,##0.00', 5, 8])).toEqual(Array(7).fill('currency'))
    expect(categories(['$#,##0.00', '\\$#,##0.00', '€#,##0.00', '#,##0.00 "₽"', '"$"0.00', '$#,##0_);($#,##0)'])).toEqual(Array(6).fill('custom'))
    expect(categories([41, 44, '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)', '_-* #,##0.00 [$€-407]_-;-* #,##0.00 [$€-407]_-;_-* "-"?? [$€-407]_-;_-@_-'])).toEqual(Array(4).fill('accounting'))
    expect(categories(['_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'])).toEqual(['custom', 'custom'])
  })

  it("dates and times: the short and long dates, the long time, and a date or time with ;@ as Format Cells writes them", () => {
    expect(categories([14, 'm/d/yyyy', 'm/d/yy', '[$-F800]dddd, mmmm dd, yyyy', '[$-409]mmmm d, yyyy;@', 'dd/mm/yyyy;@', '[$-F400]h:mm:ss AM/PM', 'h:mm;@', '[$-409]h:mm:ss AM/PM;@'])).toEqual([
      ...Array(6).fill('date'),
      ...Array(3).fill('time'),
    ] satisfies FormatCategory[])
    expect(categories([15, 20, 22, 'h:mm', 'yyyy-mm-dd', '[$-F800]dddd, mmmm d, yyyy', '[$-F400]h:mm:ss', '[$-F400]h:mm:ss AM/PM;@', '[Red]m/d/yy'])).toEqual(Array(9).fill('custom'))
  })

  it('text, special formats, and the rest Custom', () => {
    expect(categories(['@', 49, '00000', '000-00-0000', '[<=9999999]###-####;(###) ###-####', '"x"@', ';;;', '[Color10]0.00', '[>=100]0.00;[Red]0.00'])).toEqual(['text', 'text', 'special', 'special', 'special', 'custom', 'custom', 'custom', 'custom'])
    expect(numberFormatCategory(';;;')).toEqual({ category: 'custom', name: 'Custom' })
  })
})

describe('numberFormatCategory in ru-RU', () => {
  it("names a code as Excel for Windows does: en-US's numbers, percents, exponents and text, the locale's money, fractions and short date", () => {
    expect(categories(['General', '0', '0.00', '#,##0.00', '0.00;[Red]0.00', '0%', '0.00%', '0.00E+00', '@'], 'ru-RU')).toEqual([
      'general', ...Array(4).fill('number'), 'percentage', 'percentage', 'scientific', 'text',
    ] satisfies FormatCategory[])
    expect(categories(['#,##0.00 "₽"', '#,##0.00\\ "₽";[Red]\\-#,##0.00\\ "₽"', '#,##0.00 [$₽-419];-#,##0.00 [$₽-419]', '[$$-409]#,##0.00'], 'ru-RU')).toEqual(Array(4).fill('currency'))
    expect(categories(['_-* #,##0.00 "₽"_-;-* #,##0.00 "₽"_-;_-* "-"?? "₽"_-;_-@_-', '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-', '_-* #,##0.00 [$₽-419]_-;-* #,##0.00 [$₽-419]_-;_-* "-"?? [$₽-419]_-;_-@_-'], 'ru-RU')).toEqual(Array(3).fill('accounting'))
    expect(categories(['#" "?/?', '#" "?/4', 'dd/mm/yyyy', 'dd/mm/yyyy;@', '[$-F800]dddd, mmmm dd, yyyy', '[$-F400]h:mm:ss AM/PM', 'h:mm;@'], 'ru-RU')).toEqual(['fraction', 'fraction', 'date', 'date', 'date', 'time', 'time'])
  })

  it("is (все форматы) for en-US's own forms: the dollar's money, parentheses, an escaped fraction's space, m/d/yyyy written as a code", () => {
    const en = ['"$"#,##0.00', '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)', '0.00_);(0.00)', '# ?/?', 'm/d/yyyy', 'dd.mm.yyyy', '#,##0.00 ₽', '0.00;[Red]-0.00']
    expect(categories(en, 'ru-RU')).toEqual(Array(en.length).fill('custom'))
  })

  it("a built-in format by number by the locale's own code of it: 5 to 8 Денежный, 41 to 44 Финансовый, 37 to 40 none, 14 and 27 to 36 the short date", () => {
    const ids = [5, 8, 41, 44, 37, 40, 12, 14, 27, 36, 32, 22, 49]
    expect(categories(ids, 'ru-RU')).toEqual(['currency', 'currency', 'accounting', 'accounting', 'custom', 'custom', 'fraction', 'date', 'date', 'date', 'custom', 'custom', 'text'])
    expect([numberFormatCategory('0.000', ruRU), numberFormatCategory(0, ruRU), numberFormatCategory('[Red]0.00', ruRU)]).toEqual([
      { category: 'number', name: 'Числовой' },
      { category: 'general', name: 'Общий' },
      { category: 'custom', name: '(все форматы)' },
    ])
  })

  it('a locale object without rules of computation reads codes as en-US does, by en-US names', () => {
    const { decimalSeparator, listSeparator, arrayColumnSeparator, arrayRowSeparator, booleans, errors, functions } = ruRU
    const text = { tag: 'xx-XX', decimalSeparator, listSeparator, arrayColumnSeparator, arrayRowSeparator, booleans, errors, functions }
    expect([numberFormatCategory('0.000', text), numberFormatCategory('m/d/yy', text)]).toEqual([{ category: 'number', name: 'Number' }, { category: 'date', name: 'Date' }])
  })
})

describe('numberFormatChoices', () => {
  it("the list's items in order, with the format each applies and its code in the locale", () => {
    expect(numberFormatChoices()).toEqual([
      { key: 'general', name: 'General', format: 0, code: 'General' },
      { key: 'number', name: 'Number', format: 2, code: '0.00' },
      { key: 'currency', name: 'Currency', format: '"$"#,##0.00', code: '"$"#,##0.00' },
      { key: 'accounting', name: 'Accounting', format: 44, code: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)' },
      { key: 'shortDate', name: 'Short Date', format: 14, code: 'm/d/yyyy' },
      { key: 'longDate', name: 'Long Date', format: '[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy', code: '[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy' },
      { key: 'time', name: 'Time', format: '[$-F400]h:mm:ss\\ AM/PM', code: '[$-F400]h:mm:ss\\ AM/PM' },
      { key: 'percentage', name: 'Percentage', format: 10, code: '0.00%' },
      { key: 'fraction', name: 'Fraction', format: 12, code: '# ?/?' },
      { key: 'scientific', name: 'Scientific', format: 11, code: '0.00E+00' },
      { key: 'text', name: 'Text', format: 49, code: '@' },
    ])
    const ru = numberFormatChoices('ru-RU')
    expect(ru.map((c) => c.name)).toEqual(['Общий', 'Числовой', 'Денежный', 'Финансовый', 'Краткий формат даты', 'Длинный формат даты', 'Время', 'Процентный', 'Дробный', 'Экспоненциальный', 'Текстовый'])
    expect([ru[2].format, ru[3].code, ru[4].code]).toEqual(['#,##0.00\\ "₽"', '_-* #,##0.00 "₽"_-;-* #,##0.00 "₽"_-;_-* "-"?? "₽"_-;_-@_-', 'dd.mm.yyyy'])
  })

  it("each item's format is in the box's category of the item, in en-US and ru-RU", () => {
    const want: Record<string, FormatCategory> = { general: 'general', number: 'number', currency: 'currency', accounting: 'accounting', shortDate: 'date', longDate: 'date', time: 'time', percentage: 'percentage', fraction: 'fraction', scientific: 'scientific', text: 'text' }
    for (const tag of ['en-US', 'ru-RU']) expect(numberFormatChoices(tag).map((c) => numberFormatCategory(c.format, tag).category)).toEqual(Object.values(want))
  })
})

describe('input', () => {
  it('checks the format and the locale, and refuses options it does not take', () => {
    expect(() => numberFormatCategory(true as never)).toThrow(new InputError('format', 'expected a format code or the number of a built-in format, 0 to 49, got boolean true'))
    expect(() => numberFormatCategory(50)).toThrow(InputError)
    expect(() => numberFormatChoices('de-DE')).toThrow(new InputError('locale', `unknown locale "de-DE" (loaded: ${[...LOCALES.keys()].join(', ')}; the package's: en-US, ru-RU)`))
    expect(() => (numberFormatCategory as (...a: unknown[]) => unknown)('0', 'en-US', { locale: 'ru-RU' })).toThrow(InputError)
  })
})
