/**
 * Rules of computation that are a locale's data rather than code written for ru-RU: the order of day, month and year in
 * date text (`RegionalData.dateOrder`) and where the currency symbol stands (`currency.before`). ru-RU reads as Excel
 * ru-RU does (its sets in `test/excel/cases/ru-RU/`); the other orders and a symbol before the amount are the engine's
 * own rules, by analogy with en-US, and are checked here on ru-RU's data with that one field changed.
 */
import { describe, expect, it } from 'vitest'
import { serialOf } from '../src/dates.js'
import type { RegionalData } from '../src/locale/rules/data.js'
import { regionalRules } from '../src/locale/rules/regional.js'
import { ruRU as ru } from '../src/locales/ru-RU.js'

const clock = () => ({ year: () => 2026, read: false })
const rulesWith = (change: Partial<RegionalData>) => regionalRules(ru, { ...ru.regional!, ...change }, clock(), () => undefined)
const read = (change: Partial<RegionalData>, texts: string[]) => texts.map((t) => rulesWith(change).valueText(t))

describe('the order of date text', () => {
  it('ru-RU is day first, ISO dates read too', () => {
    expect(ru.regional!.dateOrder).toBe('dmy')
    expect(read({}, ['31.01.2025', '31/1/25', '2025-01-31', '2025.1.31', '1.5', '1.13', '1.2025', '31 янв 2025', 'янв 2025', '1/31/2025'])).toEqual([
      serialOf(2025, 1, 31),
      serialOf(2025, 1, 31),
      serialOf(2025, 1, 31),
      serialOf(2025, 1, 31),
      serialOf(2026, 5, 1),
      serialOf(2013, 1, 1),
      serialOf(2025, 1, 1),
      serialOf(2025, 1, 31),
      serialOf(2025, 1, 1),
      null,
    ])
  })

  it('month first: three numbers month, day, year; two a month and a day or a year, as en-US reads them', () => {
    expect(read({ dateOrder: 'mdy' }, ['1/31/2025', '1.31.25', '2025-01-31', '1/31', '1/2025', '31/1/2025', '15 янв 2025'])).toEqual([
      serialOf(2025, 1, 31),
      serialOf(2025, 1, 31),
      serialOf(2025, 1, 31),
      serialOf(2026, 1, 31),
      serialOf(2025, 1, 1),
      null,
      serialOf(2025, 1, 15),
    ])
  })

  it('year first: two digits of the year first too', () => {
    expect(read({ dateOrder: 'ymd' }, ['2025/1/31', '25.1.31', '1/31', '31.01.2025'])).toEqual([serialOf(2025, 1, 31), serialOf(2025, 1, 31), serialOf(2026, 1, 31), null])
  })
})

describe('the currency symbol in number text', () => {
  it('ru-RU reads it after the amount only', () => {
    expect(read({}, ['1 234,5 ₽', '-1 234,5 ₽', '₽1 234,5', '1 1/2 ₽', '5% ₽'])).toEqual([1234.5, -1234.5, null, 1.5, null])
  })

  it('a symbol written before the amount is read there, with a sign before or after it, as en-US reads $', () => {
    const before = { currency: { symbol: '€', before: true, space: ' ', negative: '-' as const } }
    expect(read(before, ['€ 1 234,5', '-€ 5', '€-5', '(€5)', '-€-5', '1 234,5 €', '€ 1 1/2', '€5%', '5'])).toEqual([1234.5, -5, -5, -5, null, null, null, null, 5])
  })
})
