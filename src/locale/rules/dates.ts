/**
 * The date part of date text as a day-month-year locale reads it, measured in ru-RU (`DateSyntax.date`, the
 * text in lower case): `31.01.2025`, `31/1/25`, `31-01-2025`, `2025-01-31`, `2025.1.31`; month names by a
 * prefix of three letters or more of the nominative (`31 январь 2025`, `31.янв.2025`, `31-сент-25`, `15 июн. 2025`,
 * `янв 2025`, `15 янв`). Two numbers are a day and a month (`1.5`), or a month and a year when the second is
 * past 12 (`1.13` is January 2013). A date without a year takes the clock's year, as in Excel.
 */
import { calendarDate, fullYear } from '../../dates.js'
import type { RegionalData, RulesClock } from './data.js'

const YMD = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/
const DMY = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4}|\d{2})$/
const TWO = /^(\d{1,2})[-/.](\d{1,4})$/
// after a month name a dot, a dot and a space, a space or a hyphen
const DAY_NAME = /^(\d{1,2})[ .-](\p{L}+)(?:(?:\. ?|[ -])(\d{4}|\d{2}))?$/u
const NAME_YEAR = /^(\p{L}+)(?:\. ?|[ -])(\d{4}|\d{2})$/u

export function dmyDates(data: RegionalData, clock: RulesClock): (s: string) => number | null {
  // a month by a prefix of its name, three letters at least; 0 if none
  const month = (word: string) => (word.length < 3 ? 0 : data.months.findIndex((m) => m.startsWith(word)) + 1)
  return (s) => {
    let m = YMD.exec(s)
    if (m) return calendarDate(Number(m[1]), Number(m[2]), Number(m[3]))
    if ((m = DMY.exec(s))) return calendarDate(fullYear(m[3]), Number(m[2]), Number(m[1]))
    if ((m = TWO.exec(s))) {
      const [a, b] = [Number(m[1]), m[2]]
      if (b.length <= 2 && Number(b) <= 12) return calendarDate(clock.year(), Number(b), a)
      return b.length === 3 ? null : calendarDate(fullYear(b), a, 1)
    }
    if ((m = DAY_NAME.exec(s))) {
      const mo = month(m[2])
      return mo ? calendarDate(m[3] === undefined ? clock.year() : fullYear(m[3]), mo, Number(m[1])) : null
    }
    if ((m = NAME_YEAR.exec(s))) {
      const mo = month(m[1])
      return mo ? calendarDate(fullYear(m[2]), mo, 1) : null
    }
    return null
  }
}
