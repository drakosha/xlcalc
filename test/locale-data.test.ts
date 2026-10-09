/**
 * Rules of computation that are a locale's data rather than code written for ru-RU: the order of day, month and year in
 * date text (`RegionalData.dateOrder`). ru-RU reads as Excel ru-RU does (its sets in `test/excel/cases/ru-RU/`); the other
 * orders are the engine's own rules, by analogy with en-US, and are checked here on ru-RU's data with that one field changed.
 */
import { describe, expect, it } from 'vitest'
import { serialOf } from '../src/dates.js'
import { LOCALE_FILES } from '../src/locale/registry.js'
import type { RegionalData } from '../src/locale/rules/data.js'
import { regionalRules } from '../src/locale/rules/regional.js'

const ru = LOCALE_FILES.get('ru-RU')!
const clock = () => ({ year: () => 2026, read: false })
const rulesWith = (change: Partial<RegionalData>) => regionalRules(ru.text, { ...ru.regional!, ...change }, clock(), () => undefined)
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
