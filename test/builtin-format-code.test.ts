/**
 * `builtinFormatCode`: the code a built-in number format shows by in a locale, the one the engine shows the number by
 * (the same text for every value), undefined for a number that is no built-in format's; its input checks.
 */
import { describe, expect, it } from 'vitest'
import { builtinFormatCode, displayValue, InputError, LOCALES } from '../src/index.js'

const VALUES = [0, 1234.567, -1234.5, 0.25, 46096.75, -0.5, 'text']

describe('builtinFormatCode', () => {
  it("gives the code Excel shows a built-in format by in the locale: 14 is the locale's short date", () => {
    expect([builtinFormatCode(0), builtinFormatCode(14), builtinFormatCode(14, 'en-US'), builtinFormatCode(14, 'ru-RU')]).toEqual(['General', 'm/d/yyyy', 'm/d/yyyy', 'dd.mm.yyyy'])
    expect([builtinFormatCode(7, 'ru-RU'), builtinFormatCode(22, 'ru-RU'), builtinFormatCode(49, 'ru-RU')]).toEqual(['#,##0.00 "₽";-#,##0.00 "₽"', 'dd.mm.yyyy h:mm', '@'])
    expect(builtinFormatCode(9, 'ru-RU')).toBe('0%')
  })

  it('undefined for a number that is no built-in format: a built-in format is one it has a code for', () => {
    for (const id of [-1, 50, 164, 1.5, Number.NaN, Infinity]) expect(builtinFormatCode(id), String(id)).toBeUndefined()
    expect(Array.from({ length: 60 }, (_, id) => builtinFormatCode(id) !== undefined).lastIndexOf(true)).toBe(49)
  })

  it('a value shows by the code as by the number, in every locale and for every built-in format', () => {
    const off: string[] = []
    for (const locale of LOCALES.keys())
      for (let id = 0; id < 50; id++)
        for (const v of VALUES) {
          const [byId, byCode] = [displayValue(v, { format: id, locale }), displayValue(v, { format: builtinFormatCode(id, locale), locale })]
          if (JSON.stringify(byId) !== JSON.stringify(byCode)) off.push(`${locale} ${id} ${v}: ${JSON.stringify(byId)} by the number, ${JSON.stringify(byCode)} by the code`)
        }
    expect(off).toEqual([])
  })

  it('checks its input', () => {
    expect(() => builtinFormatCode('14' as never)).toThrow(new InputError('id', 'expected the number of a built-in format, 0 to 49, got the string "14"'))
    expect(() => builtinFormatCode(14, 'de-DE')).toThrow(new InputError('locale', `unknown locale "de-DE" (expected ${[...LOCALES.keys()].join(', ')})`))
  })
})
