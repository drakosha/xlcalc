/**
 * The date part of date text as a locale other than en-US reads it (`DateSyntax.date`, the text in lower case), by its
 * order of day, month and year (`RegionalData.dateOrder`). Day-month-year was measured in ru-RU: `31.01.2025`, `31/1/25`,
 * `31-01-2025`, `1-2-3` (a year of one digit, the entry set `typed`), also `2025-01-31` and `2025.1.31`; month names by a prefix of three letters or more of the nominative
 * (`31 январь 2025`, `31.янв.2025`, `31-сент-25`, `15 июн. 2025`, `янв 2025`, `15 янв`). Two numbers are a day and a month
 * (`1.5`), or a month and a year when the second is past 12 (`1.13` is January 2013). A date without a year takes the
 * clock's year, as in Excel. In month-day-year and year-month-day text the three numbers take their places by the order,
 * and two numbers are a month and then a day or a year, as en-US reads `1/31` and `1/2025`; no Excel of such a locale
 * has been measured, and month names are read day first in every order.
 */
import { calendarDate, fullYear, monthAnd } from '../../dates.js'
import type { DateOrder, RegionalData, RulesClock } from './data.js'

const YMD = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/
const TWO = /^(\d{1,2})[-/.](\d{1,4})$/
// after a month name a dot, a dot and a space, a space or a hyphen
const DAY_NAME = /^(\d{1,2})[ .-](\p{L}+)(?:(?:\. ?|[ -])(\d{4}|\d{2}))?$/u
const NAME_YEAR = /^(\p{L}+)(?:\. ?|[ -])(\d{4}|\d{2})$/u

/** Three numbers in each order: the pattern, and the places of the year, the month and the day in it. */
const THREE: Record<DateOrder, { re: RegExp; y: number; m: number; d: number }> = {
  dmy: { re: /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4}|\d{1,2})$/, y: 3, m: 2, d: 1 },
  mdy: { re: /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4}|\d{1,2})$/, y: 3, m: 1, d: 2 },
  ymd: { re: /^(\d{4}|\d{2})[-/.](\d{1,2})[-/.](\d{1,2})$/, y: 1, m: 2, d: 3 },
}

export function regionalDates(data: RegionalData, clock: RulesClock): (s: string) => number | null {
  // a month by a prefix of its name, three letters at least, in any case (the text comes in lower case); 0 if none
  const names = data.months.map((m) => m.toLowerCase())
  const month = (word: string) => (word.length < 3 ? 0 : names.findIndex((m) => m.startsWith(word)) + 1)
  const three = THREE[data.dateOrder]
  const year = () => clock.year()
  // two numbers, `a` and then `b`: day first, a day and a month, or a month and a year past 12; else a month first
  const two =
    data.dateOrder === 'dmy'
      ? (a: number, b: string) => {
          if (b.length <= 2 && Number(b) <= 12) return calendarDate(year(), Number(b), a)
          return b.length === 3 ? null : calendarDate(fullYear(b), a, 1)
        }
      : (a: number, b: string) => monthAnd(a, b, year)
  return (s) => {
    let m = YMD.exec(s)
    if (m) return calendarDate(Number(m[1]), Number(m[2]), Number(m[3]))
    if ((m = three.re.exec(s))) return calendarDate(fullYear(m[three.y]), Number(m[three.m]), Number(m[three.d]))
    if ((m = TWO.exec(s))) return two(Number(m[1]), m[2])
    if ((m = DAY_NAME.exec(s))) {
      const mo = month(m[2])
      return mo ? calendarDate(m[3] === undefined ? year() : fullYear(m[3]), mo, Number(m[1])) : null
    }
    if ((m = NAME_YEAR.exec(s))) {
      const mo = month(m[1])
      return mo ? calendarDate(fullYear(m[2]), mo, 1) : null
    }
    return null
  }
}
