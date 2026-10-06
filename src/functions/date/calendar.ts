/** Calendar functions in Excel's 1900 system (serial 0 is 1900-01-00, 60 the phantom 1900-02-29). */
import type { Cx, FnDef } from '../../args.js'
import { addMonths, dayAt, daysInMonth, MAX_SERIAL, serialOf, ymd } from '../../dates.js'
import type { Rules } from '../../locale/rules.js'
import { E, isErr, WHOLE, type CErr, type V } from '../../values.js'
import { atp } from './atp.js'
import { lastDayExcel } from './daycount.js'

/** 0 is Sunday: Excel takes serial 1 (1900-01-01) for a Sunday. */
export const weekday = (serial: number) => (serial + 6) % 7

const serialExcel = (y: number, m: number, d: number) => (y === 1900 && m === 2 && d === 29 ? 60 : serialOf(y, m, d))

/** DATEVALUE and TIMEVALUE take text only, read by the book's locale. */
export function textDate(v: V, rules: Rules): number | CErr {
  if (typeof v !== 'string') return E.VALUE
  return rules.dateText(v) ?? E.VALUE
}

/** The week (from Monday) belongs to the year of its Thursday. */
function isoWeek(serial: number): number {
  const thursday = serial - ((weekday(serial) + 6) % 7) + 3
  return Math.floor((thursday - serialOf(ymd(thursday).y, 1, 1)) / 7) + 1
}

/** First day of the week by WEEKDAY/WEEKNUM return type. */
const WEEK_START: Record<number, number> = { 1: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0 }

/** MD and YD count from the start's day as Excel does (MD may be negative). */
function dateDif(start: number, end: number, unit: string): number | CErr {
  if (start > end) return E.NUM
  const a = ymd(start)
  const b = ymd(end)
  const months = (b.y - a.y) * 12 + (b.m - a.m) - (b.d < a.d ? 1 : 0)
  switch (unit.toUpperCase()) {
    case 'D':
      return end - start
    case 'M':
      return months
    case 'Y':
      return Math.floor(months / 12)
    case 'YM':
      return months % 12
    case 'MD': {
      if (b.d >= a.d) return b.d - a.d
      const prev = addMonths(b.y, b.m, -1)
      return b.d - a.d + lastDayExcel(prev.y, prev.m)
    }
    case 'YD':
      return end - serialExcel(b.m > a.m || (b.m === a.m && b.d >= a.d) ? b.y : b.y - 1, a.m, a.d)
    default:
      return E.NUM
  }
}

export const functions: Record<string, FnDef> = {
  // years 0–1899 are added to 1900; month and day roll over, so DATE(1900,2,29) is 60; a day below zero
  // rounds down, after the whole-number tolerance (DATE(2025,1,-1.9999997) is December 29)
  DATE: {
    args: 'w w n',
    fn: (y: number, m: number, day: number) => {
      const d = Math.floor(day + WHOLE)
      const year = y >= 0 && y < 1900 ? y + 1900 : y
      if (year < 1900 || year > 9999) return E.NUM
      const first = addMonths(year, m, 0)
      if (first.y < 1899 || first.y > 9999) return E.NUM
      const serial = serialOf(first.y, first.m, 1) + d - 1
      return serial < 0 || serial > MAX_SERIAL ? E.NUM : serial
    },
  },
  DATEDIF: { args: 'd d s', fn: dateDif },
  DATEVALUE: {
    args: 'v',
    cx: true,
    fn: (cx: Cx, v: V) => {
      const t = textDate(v, cx.rules)
      return isErr(t) ? t : Math.floor(t)
    },
  },
  DAY: { args: 't', fn: (t: number) => ymd(dayAt(t)).d },
  DAYS: { args: 'd d', fn: (end: number, start: number) => end - start },
  // the day clamped to the target month; serial 0 is day 0 of January 1900
  EDATE: atp('e j', (start: number, months: number) => {
    const s = ymd(start)
    const t = addMonths(s.y, s.m, months)
    if (t.y < 1900 || t.y > 9999) return E.NUM
    const serial = serialOf(t.y, t.m, 1) + Math.min(s.d, daysInMonth(t.y, t.m)) - 1
    return serial < 0 || serial > MAX_SERIAL ? E.NUM : serial
  }),
  // the start month in Excel's calendar (serial 60 is February 1900), the last day by the real one
  EOMONTH: atp('e j', (start: number, months: number) => {
    const s = ymd(start)
    const t = addMonths(s.y, s.m, months)
    if (t.y < 1900 || t.y > 9999) return E.NUM
    return serialOf(t.y, t.m, daysInMonth(t.y, t.m))
  }),
  ISOWEEKNUM: { args: 'd', fn: isoWeek },
  MONTH: { args: 't', fn: (t: number) => ymd(dayAt(t)).m },
  WEEKDAY: {
    args: 't w?',
    fn: (t: number, type = 1) => {
      const w = weekday(dayAt(t))
      if (type === 1) return w + 1
      if (type === 2) return ((w + 6) % 7) + 1
      if (type === 3) return (w + 6) % 7
      if (type >= 11 && type <= 17) return ((w - WEEK_START[type] + 7) % 7) + 1
      return E.NUM
    },
  },
  // week 1 contains January 1 (type 21: ISO weeks), so serial 0 is in week 0
  WEEKNUM: atp('e j?', (s: number, type = 1) => {
    if (type === 21) return isoWeek(s)
    const first = WEEK_START[type]
    if (first === undefined) return E.NUM
    const jan1 = serialOf(ymd(s).y, 1, 1)
    return Math.floor((s - jan1 + ((weekday(jan1) - first + 7) % 7)) / 7) + 1
  }),
  YEAR: { args: 't', fn: (t: number) => ymd(dayAt(t)).y },
}
