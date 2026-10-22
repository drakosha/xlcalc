/**
 * Reproductions from the review of 21.10 on number formats: digits past the 15th significant one show as zeros below
 * 1E+21 too, as they do above. In TEXT with closures and eval, and in a cell's displayed text.
 */
import { describe, expect, it } from 'vitest'
import { compute, displayValue, type BookInput } from '../src/index.js'
import './locales.js'

/** The values of the formulas, one per cell of column B. */
function values(formulas: string[], execution: 'closures' | 'eval', locale?: string) {
  const book: BookInput = { ...(locale ? { locale } : {}), sheets: [{ name: 'S', cells: Object.fromEntries(formulas.map((f, i) => ['B' + (i + 1), { formula: f }])) }] }
  const r = compute(book, { execution })
  return formulas.map((_, i) => r.sheets[0].cells['B' + (i + 1)].value)
}

describe.each(['closures', 'eval'] as const)('15 significant digits in digit formats, %s', (execution) => {
  it('writes the digits past the 15th as zeros below 1E+21', () => {
    const cases: [string, string][] = [
      ['TEXT(1/3,"0.00000000000000000000")', '0.33333333333333300000'],
      ['TEXT(2^60,"0")', '1152921504606850000'],
      ['TEXT(2^60,"#,##0")', '1,152,921,504,606,850,000'],
      ['TEXT(793853.646,"0.0000000000")', '793853.6460000000'],
      ['TEXT(86960145761258.9,"0.00")', '86960145761258.90'],
      ['TEXT(77017111.9,"0.00000000")', '77017111.90000000'],
      ['TEXT(0.1+0.2,"0.00000000000000000")', '0.30000000000000000'],
      ['TEXT(123456789012345000,"0%")', '12345678901234500000%'],
      ['TEXT(1/3*1E6,"#,##0.0000000000")', '333,333.3333333330'],
      ['TEXT(-(2^60),"0")', '-1152921504606850000'],
      ['TEXT(2^60,"# ?/?")', '1152921504606850000    '],
      ['TEXT(1E15-0.1,"0.0")', '1000000000000000.0'],
    ]
    expect(values(cases.map((c) => c[0]), execution)).toEqual(cases.map((c) => c[1]))
    expect(values(['TEXT(2^60,"# ##0")'], execution, 'ru-RU')).toEqual(['1 152 921 504 606 850 000'])
  })

  it('keeps what shows 15 significant digits or fewer', () => {
    const cases: [string, string][] = [
      ['TEXT(1.005,"0.00")', '1.01'],
      ['TEXT(1.2344999999999999,"0.000")', '1.234'],
      ['TEXT(1234.5678,"#,##0.00")', '1,234.57'],
      ['TEXT(0.000123456789012345,"0.000000000000000000")', '0.000123456789012345'],
      ['TEXT(1E+21,"0")', '1000000000000000000000'],
    ]
    expect(values(cases.map((c) => c[0]), execution)).toEqual(cases.map((c) => c[1]))
  })
})

it('displays a cell the same as TEXT writes it', () => {
  expect(displayValue(1152921504606850000, { numberFormat: '#,##0' }).text).toBe('1,152,921,504,606,850,000')
  expect(displayValue(0.333333333333333, { numberFormat: '0.' + '0'.repeat(20), width: 255 }).text).toBe('0.33333333333333300000')
  expect(displayValue(1 / 3 * 1e6, { numberFormat: '#,##0.0000000000', width: 255 }).text).toBe('333,333.3333333330')
})

describe.each(['closures', 'eval'] as const)('ru-RU TEXT with a colour name, %s', (execution) => {
  const value = { error: 'VALUE' }

  it('refuses a point before a digit placeholder whatever the colour, as without one', () => {
    const colours = ['Красный', 'Синий', 'Черный', 'Голубой', 'Зеленый', 'Желтый', 'Белый', 'Цвет10']
    const formulas = ['TEXT(1234.5,"0.00")', ...colours.map((c) => `TEXT(1234.5,"[${c}]0.00")`), 'TEXT(1234.5,"[Красный]#.##")']
    expect(values(formulas, execution, 'ru-RU')).toEqual(formulas.map(() => value))
  })

  it('keeps colours on numbers, dates and elapsed time', () => {
    const cases: [string, string][] = [
      ['TEXT(1234.567,"[Красный]0")', '1235'],
      ['TEXT(1234.5,"[Красный]0,00")', '1234,50'],
      ['TEXT(1234.5,"[Синий]# ##0,00")', '1 234,50'],
      ['TEXT(45688,"[Красный]ДД.ММ.ГГГГ")', '31.01.2025'],
      ['TEXT(1.25,"[ч]:мм:сс")', '30:00:00'],
      ['TEXT(1.25,"[Черный][ч]:мм")', '30:00'],
    ]
    expect(values(cases.map((c) => c[0]), execution, 'ru-RU')).toEqual(cases.map((c) => c[1]))
  })
})

describe.each(['closures', 'eval'] as const)('a code ending in a text section, %s', (execution) => {
  const value = { error: 'VALUE' }

  it('gives numbers the sections before the text section, as the code without it', () => {
    const cases: [string, unknown][] = [
      ['TEXT(-0.25,"h:mm;@")', value],
      ['TEXT(-1,"m/d/yyyy;@")', value],
      ['TEXT(-0.25,"[$-409]h:mm AM/PM;@")', value],
      ['TEXT(-5,"0;@")', '-5'],
      ['TEXT(-5,"0.00;""t""@")', '-5.00'],
      ['TEXT(0,"0.00;-0.00;@")', '0.00'],
      ['TEXT(-5,"[>10]0;@")', value],
      ['TEXT(5,"[<0]0;[>10]0;@")', value],
      // the one number section before the text section writes the minus, as Excel answered (the set values-held)
      ['TEXT(-5,"[<0]0;@")', '-5'],
    ]
    expect(values(cases.map((c) => c[0]), execution)).toEqual(cases.map((c) => c[1]))
  })

  it('keeps what Excel already gave', () => {
    const cases: [string, unknown][] = [
      ['TEXT(1.5,"h:mm;@")', '12:00'],
      ['TEXT("a","h:mm;@")', 'a'],
      ['TEXT(-5,"0;0;@")', '5'],
      ['TEXT(0,"0.00;@")', '0.00'],
      ['TEXT(0,"0;@")', '0'],
      ['TEXT(-5,"0.00;-0.00;@")', '-5.00'],
      ['TEXT(-1235,"0;-0;;@")', '-1235'],
      ['TEXT(-5,"[<0]0;[>10]0;@")', '5'],
      ['TEXT(20,"[<0]0;[>10]0;@")', '20'],
      ['TEXT(-5,"@")', '-5'],
    ]
    expect(values(cases.map((c) => c[0]), execution)).toEqual(cases.map((c) => c[1]))
  })

  it('displays a negative date or time under it as # in a cell', () => {
    const r = compute({ sheets: [{ name: 'S', cells: { A1: { value: -0.25, numberFormat: 'h:mm;@' }, A2: { value: -5, numberFormat: '0;@' } } }] }, { execution })
    expect(r.sheets[0].cells.A1.text).toMatch(/^#+$/)
    expect(r.sheets[0].cells.A2.text).toBe('-5')
    expect(displayValue(-1, { numberFormat: 'm/d/yy;@' }).text).toMatch(/^#+$/)
  })
})
