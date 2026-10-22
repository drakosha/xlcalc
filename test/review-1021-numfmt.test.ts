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
