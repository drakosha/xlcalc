/**
 * Excel's 1900 date system and date and time text in the en-US locale (`src/dates.ts`).
 */
import { describe, expect, it } from 'vitest'
import { parseDateText, serialOf, ymd } from '../src/dates.js'

describe('1900 date system', () => {
  it('serial numbers and calendar dates, with day 0 and the phantom 1900-02-29', () => {
    expect([serialOf(1900, 1, 1), serialOf(1900, 2, 28), serialOf(1900, 3, 1), serialOf(2025, 1, 31), serialOf(9999, 12, 31)]).toEqual([1, 59, 61, 45688, 2958465])
    expect([ymd(0), ymd(59), ymd(60), ymd(61), ymd(45688.75)]).toEqual([
      { y: 1900, m: 1, d: 0 },
      { y: 1900, m: 2, d: 28 },
      { y: 1900, m: 2, d: 29 },
      { y: 1900, m: 3, d: 1 },
      { y: 2025, m: 1, d: 31 },
    ])
  })

  it('date and time text as Excel en-US reads it', () => {
    const forms = ['2025-01-31', '2025/1/31', '1/31/2025', '1-31-2025', '1 / 31 / 2025', '1/31/25', '31-Jan-2025', '31 january 2025', 'Jan 31, 2025', 'Janu 31, 2025']
    for (const f of forms) expect(parseDateText(f), f).toBe(45688)
    expect(parseDateText('1/31/30')).toBe(serialOf(1930, 1, 31))
    expect(parseDateText('2/29/1900')).toBe(60)
    expect(parseDateText('18:30')).toBeCloseTo(18.5 / 24, 12)
    expect(parseDateText('6:30:15 pm')).toBeCloseTo((18 * 3600 + 30 * 60 + 15) / 86400, 12)
    expect(parseDateText('12 AM')).toBe(0)
    expect(parseDateText('1/31/2025 6 PM')).toBe(45688.75)
    expect(parseDateText('18:00 1/31/2025')).toBe(45688.75)
    expect(parseDateText('1/2025')).toBe(45658)
    // past 23 hours and 59 minutes without AM/PM, as Excel en-US reads them
    expect(parseDateText('25:00')).toBeCloseTo(25 / 24, 12)
    expect(parseDateText('10:60')).toBeCloseTo(11 / 24, 12)
    expect(parseDateText('0 PM')).toBe(0.5)
  })

  it('what is not a date: impossible dates and times, no year, other text', () => {
    // `T` between the date and the time, a year without a comma, a space or a tab before the text: no date in Excel (`time-text`, `date-noyear`, `lead-spaces` sets)
    for (const f of ['2/29/2025', '13/1/2025', '1/31', '13:00 PM', '12:60 PM', '10000:00', 'abc', '', '1899-12-31', '45688', '2025-01-31T18:00', 'January 31 2025', ' 1/31/2025', '\t6:30', '6pm'])
      expect(parseDateText(f), f).toBeNull()
  })
})
