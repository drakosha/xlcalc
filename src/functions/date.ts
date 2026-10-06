/**
 * Date and time functions: this engine's own, in Excel's 1900 date system (`src/dates.ts`): serial
 * numbers, day 0 (1900-01-00) and the phantom 1900-02-29 (serial 60). Weekdays are serial arithmetic
 * (Excel takes serial 1 for a Sunday), never JS Date in the local zone. Also the day-count conventions
 * (30/360, actual/actual and the rest) the financial functions share, and the argument rule of the former
 * Analysis ToolPak functions (`atp`).
 */
import { toDate, type Cx, type FnDef } from '../args.js'
import { addMonths, daysInMonth, isLeap, MAX_SERIAL, parseDateText, serialOf, ymd } from '../dates.js'
import { E, isErr, toBool, toNum, toText, type CErr, type V } from '../values.js'

// ---------- the argument rule of the former Analysis ToolPak ----------

const noBool = (v: V): number | CErr => (typeof v === 'boolean' ? E.VALUE : toNum(v))

/** Scalar kinds `atp` coerces itself, in argument order. */
const SCALAR: Record<string, (v: V) => unknown> = {
  e: (v) => (typeof v === 'boolean' ? E.VALUE : toDate(v)),
  x: noBool,
  j: (v) => {
    const n = noBool(v)
    return isErr(n) ? n : Math.trunc(n)
  },
  n: toNum,
  i: (v) => {
    const n = toNum(v)
    return isErr(n) ? n : Math.trunc(n)
  },
  d: toDate,
  b: toBool,
  s: toText,
}

/**
 * A function of the former Analysis ToolPak (bond and coupon functions, EDATE, WORKDAY and the like): a
 * boolean is #VALUE! where a number or a date is expected. Kinds: `e` a date (date text read, a boolean
 * #VALUE!), `x` a number and `j` an integer (truncated) with a boolean #VALUE!, `n`, `i`, `d`, `b`, `s` as
 * in `args.ts`. These are coerced here in argument order, the first error being the answer; other kinds
 * (lists, `a`) go through the shared layer.
 */
export function atp(kinds: string, fn: (...a: any[]) => V): FnDef {
  const ks = kinds.split(' ').filter(Boolean)
  const local = ks.map((k) => k[0] in SCALAR)
  return {
    args: ks.map((k, i) => (local[i] ? `a${k.slice(1)}` : k)).join(' '),
    fn: (...a: unknown[]) => {
      for (let i = 0; i < a.length; i++) {
        if (!local[i] || a[i] === undefined) continue
        const v = a[i] as V
        if (isErr(v)) return v
        const c = SCALAR[ks[i][0]](v)
        if (isErr(c)) return c
        a[i] = c
      }
      return fn(...a)
    },
  }
}

// ---------- calendar helpers ----------

/** Day of the week, 0 for Sunday: Excel takes serial 1 (1900-01-01) for a Sunday and serial 0 for a Saturday. */
export const weekday = (serial: number) => (serial + 6) % 7

/** The last day of a month in Excel's calendar, where February 1900 has 29 days. */
const lastDayExcel = (y: number, m: number) => (y === 1900 && m === 2 ? 29 : daysInMonth(y, m))

/** The serial number of a calendar date with Excel's 1900-02-29; day 0 is the day before the 1st. */
const serialExcel = (y: number, m: number, d: number) => (y === 1900 && m === 2 && d === 29 ? 60 : serialOf(y, m, d))

const inRange = (serial: number): number | CErr => (serial < 0 || serial > MAX_SERIAL ? E.NUM : serial)

/** Hours, minutes and seconds of a serial number's day fraction, rounded to the nearest second. */
function clock(t: number): { h: number; m: number; s: number } {
  let secs = Math.round((t - Math.floor(t)) * 86_400)
  if (secs === 86_400) secs = 0
  return { h: Math.floor(secs / 3600), m: Math.floor(secs / 60) % 60, s: secs % 60 }
}

/** A date or date-time text as DATEVALUE and TIMEVALUE read it: a number or other value is #VALUE!. */
function textDate(v: V): number | CErr {
  if (typeof v !== 'string') return E.VALUE
  return parseDateText(v) ?? E.VALUE
}

// ---------- day-count conventions ----------

/**
 * Days between two dates on a 30-day-month calendar, as DAYS360 counts them: the US (NASD) method moves a
 * start on the 31st or on the last day of February (Excel's calendar: 1900-02-29) to the 30th, and an end
 * on the 31st to the 30th when the start is on the 30th or later; the European method moves both 31sts.
 */
export function days360(start: number, end: number, european: boolean): number {
  const a = ymd(start)
  const b = ymd(end)
  let d1 = a.d
  let d2 = b.d
  if (european) {
    d1 = Math.min(d1, 30)
    d2 = Math.min(d2, 30)
  } else {
    if (d1 === 31 || (a.m === 2 && d1 === lastDayExcel(a.y, 2))) d1 = 30
    if (d2 === 31 && d1 >= 30) d2 = 30
  }
  return (b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1)
}

/** Whether a date is the last day of February by the real calendar (1900 is not a leap year here). */
const lastOfFeb = (x: { y: number; m: number; d: number }) => x.m === 2 && x.d === (isLeap(x.y) ? 29 : 28)

/** YEARFRAC's 30/360 (basis 0) between ordered dates: NASD with the February rules of the Analysis ToolPak. */
function yearFrac0(start: number, end: number): number {
  const a = ymd(start)
  const b = ymd(end)
  let d1 = a.d
  let d2 = b.d
  if (d1 === 31 && d2 === 31) {
    d1 = 30
    d2 = 30
  } else if (d1 === 31) d1 = 30
  else if (d1 === 30 && d2 === 31) d2 = 30
  else if (lastOfFeb(a) && lastOfFeb(b)) {
    d1 = 30
    d2 = 30
  } else if (lastOfFeb(a)) d1 = 30
  return ((b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1)) / 360
}

/** YEARFRAC's actual/actual (basis 1) between ordered dates: one year's length, or the average over the years spanned. */
function yearFrac1(start: number, end: number): number {
  const a = ymd(start)
  const b = ymd(end)
  const days = end - start
  const withinYear = a.y === b.y || (b.y === a.y + 1 && (a.m > b.m || (a.m === b.m && a.d >= b.d)))
  if (withinYear) {
    let leap = false
    if (a.y === b.y) leap = isLeap(a.y)
    else {
      const mar1a = serialOf(a.y, 3, 1)
      const mar1b = serialOf(b.y, 3, 1)
      leap = (isLeap(a.y) && start < mar1a && end >= mar1a) || (isLeap(b.y) && end >= mar1b && start < mar1b)
    }
    if (b.m === 2 && b.d === 29) leap = true
    return days / (leap ? 366 : 365)
  }
  const years = b.y - a.y + 1
  return days / ((serialOf(b.y + 1, 1, 1) - serialOf(a.y, 1, 1)) / years)
}

/** The fraction of a year between two dates by a day-count basis (0 US 30/360, 1 actual/actual, 2 actual/360, 3 actual/365, 4 European 30/360); the dates in either order. */
export function yearFrac(start: number, end: number, basis: number): number {
  if (start > end) [start, end] = [end, start]
  switch (basis) {
    case 0:
      return yearFrac0(start, end)
    case 1:
      return yearFrac1(start, end)
    case 2:
      return (end - start) / 360
    case 3:
      return (end - start) / 365
    default:
      return days360(start, end, true) / 360
  }
}

export const validBasis = (b: number) => b >= 0 && b <= 4

// ---------- workdays ----------

/** Weekend days by WORKDAY.INTL's code (1–7 two days, 11–17 one day) or a string of seven 0/1 from Monday; index 0 is Sunday. */
function weekendOf(w: V | undefined): boolean[] | CErr {
  const mask = [false, false, false, false, false, false, false]
  if (w === undefined || w === null) w = 1
  if (isErr(w)) return E.NUM
  if (typeof w === 'boolean') return E.VALUE
  if (typeof w === 'string') {
    if (!/^[01]{7}$/.test(w) || w === '1111111') return E.VALUE
    for (let i = 0; i < 7; i++) mask[(i + 1) % 7] = w[i] === '1'
    return mask
  }
  const k = Math.trunc(w)
  if (k >= 1 && k <= 7) {
    mask[(k + 5) % 7] = true
    mask[(k + 6) % 7] = true
  } else if (k >= 11 && k <= 17) mask[k - 11] = true
  else return E.NUM
  return mask
}

/** Whole days as a set (holidays), kept apart from the weekend. */
const holidaySet = (days: number[] | undefined) => new Set(days ?? [])

/** Workdays from `start` to `end` inclusive, negative when `end` is earlier. */
function networkDays(start: number, end: number, weekend: boolean[], holidays: Set<number>): number {
  const sign = start <= end ? 1 : -1
  const [a, b] = sign > 0 ? [start, end] : [end, start]
  const perWeek = weekend.filter((x) => !x).length
  const span = b - a + 1
  let count = Math.floor(span / 7) * perWeek
  for (let s = a + Math.floor(span / 7) * 7; s <= b; s++) if (!weekend[weekday(s)]) count++
  for (const h of holidays) if (h >= a && h <= b && !weekend[weekday(h)]) count--
  return sign * count
}

/** The date `days` workdays from `start` (back for a negative count); `start` itself when `days` is 0. */
function workday(start: number, days: number, weekend: boolean[], holidays: Set<number>): number | CErr {
  if (days === 0) return start
  if (Math.abs(days) > MAX_SERIAL) return E.NUM
  const step = days > 0 ? 1 : -1
  const perWeek = weekend.filter((x) => !x).length
  let left = Math.abs(days)
  // whole weeks first (each has perWeek workdays wherever it starts), then day by day, ignoring holidays
  const weeks = Math.floor((left - 1) / perWeek)
  let d = start + step * weeks * 7
  left -= weeks * perWeek
  while (left > 0) {
    d += step
    if (!weekend[weekday(d)]) left--
  }
  // holidays passed on workdays are made up past the date reached
  let extra = 0
  for (const h of holidays) if ((step > 0 ? h > start && h <= d : h < start && h >= d) && !weekend[weekday(h)]) extra++
  while (extra > 0) {
    d += step
    if (d < 0 || d > MAX_SERIAL) return E.NUM
    if (!weekend[weekday(d)] && !holidays.has(d)) extra--
  }
  return inRange(d)
}

// ---------- weeks ----------

/** ISO 8601 week number: the week (from Monday) belongs to the year of its Thursday. */
function isoWeek(serial: number): number {
  const thursday = serial - ((weekday(serial) + 6) % 7) + 3
  const y = thursday <= 0 ? 1899 : ymd(thursday).y
  return Math.floor((thursday - serialOf(y, 1, 1)) / 7) + 1
}

/** WEEKNUM's first day of the week by return type: 1 and 17 Sunday, 2 and 11 Monday, 12–16 Tuesday–Saturday. */
const WEEK_START: Record<number, number> = { 1: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0 }

// ---------- DATEDIF ----------

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
    case 'YD': {
      const y = b.m > a.m || (b.m === a.m && b.d >= a.d) ? b.y : b.y - 1
      return end - serialExcel(y, a.m, a.d)
    }
    default:
      return E.NUM
  }
}

export const functions: Record<string, FnDef> = {
  // years 0–1899 are added to 1900; month and day roll over through Excel's calendar, so DATE(1900,2,29) is 60
  DATE: {
    args: 'i i i',
    fn: (y: number, m: number, d: number) => {
      const year = y >= 0 && y < 1900 ? y + 1900 : y
      if (year < 1900 || year > 9999) return E.NUM
      const first = addMonths(year, m, 0)
      if (first.y < 1899 || first.y > 9999) return E.NUM
      const serial = serialOf(first.y, first.m, 1) + d - 1
      return serial < 0 || serial > MAX_SERIAL ? E.NUM : serial
    },
  },
  // whole units between two dates; MD and YD count from the day of the start as Excel does (MD may be negative)
  DATEDIF: { args: 'd d s', fn: dateDif },
  DATEVALUE: {
    args: 'v',
    fn: (v: V) => {
      const t = textDate(v)
      return isErr(t) ? t : Math.floor(t)
    },
  },
  DAY: { args: 'd', fn: (s: number) => ymd(s).d },
  DAYS: { args: 'd d', fn: (end: number, start: number) => end - start },
  DAYS360: { args: 'd d b?', fn: (start: number, end: number, european = false) => days360(start, end, european) },
  // the day clamped to the target month's length; serial 0 counts as day 0 of January 1900
  EDATE: atp('e j', (start: number, months: number) => {
    const s = ymd(start)
    const t = addMonths(s.y, s.m, months)
    if (t.y < 1900 || t.y > 9999) return E.NUM
    return inRange(serialOf(t.y, t.m, 1) + Math.min(s.d, daysInMonth(t.y, t.m)) - 1)
  }),
  // the start month in Excel's calendar (serial 0 is January 1900, 60 is February); the last day by the real one
  EOMONTH: {
    args: 'd i',
    fn: (start: number, months: number) => {
      const s = ymd(start)
      const t = addMonths(s.y, s.m, months)
      if (t.y < 1900 || t.y > 9999) return E.NUM
      return serialOf(t.y, t.m, daysInMonth(t.y, t.m))
    },
  },
  HOUR: { args: 't', fn: (t: number) => clock(t).h },
  ISOWEEKNUM: { args: 'd', fn: isoWeek },
  MINUTE: { args: 't', fn: (t: number) => clock(t).m },
  MONTH: { args: 'd', fn: (s: number) => ymd(s).m },
  NETWORKDAYS: atp('e e D?', (start: number, end: number, hol?: number[]) => networkDays(start, end, weekendOf(1) as boolean[], holidaySet(hol))),
  'NETWORKDAYS.INTL': atp('e e a? D?', (start: number, end: number, w: V | undefined, hol?: number[]) => {
    const weekend = weekendOf(w)
    return isErr(weekend) ? weekend : networkDays(start, end, weekend, holidaySet(hol))
  }),
  NOW: { args: '', cx: true, volatile: true, fn: (cx: Cx) => cx.now },
  SECOND: { args: 't', fn: (t: number) => clock(t).s },
  // whole hours, minutes and seconds (INT of each, up to 32767), wrapped to a day; a negative total is #NUM!
  TIME: {
    args: 'n n n',
    fn: (h: number, m: number, s: number) => {
      const [hh, mm, ss] = [Math.floor(h), Math.floor(m), Math.floor(s)]
      if (hh > 32767 || mm > 32767 || ss > 32767) return E.NUM
      const total = hh * 3600 + mm * 60 + ss
      return total < 0 ? E.NUM : (total % 86_400) / 86_400
    },
  },
  TIMEVALUE: {
    args: 'v',
    fn: (v: V) => {
      const t = textDate(v)
      return isErr(t) ? t : t - Math.floor(t)
    },
  },
  TODAY: { args: '', cx: true, volatile: true, fn: (cx: Cx) => Math.floor(cx.now) },
  // return types 1–3 and 11–17
  WEEKDAY: {
    args: 'd i?',
    fn: (s: number, type = 1) => {
      const w = weekday(s)
      if (type === 1) return w + 1
      if (type === 2) return ((w + 6) % 7) + 1
      if (type === 3) return (w + 6) % 7
      if (type >= 11 && type <= 17) return ((w - WEEK_START[type] + 7) % 7) + 1
      return E.NUM
    },
  },
  // the week containing January 1 is week 1 (return type 21: ISO weeks); serial 0 is in week 0
  WEEKNUM: atp('e j?', (s: number, type = 1) => {
    if (type === 21) return isoWeek(s)
    const first = WEEK_START[type]
    if (first === undefined) return E.NUM
    const y = ymd(s).y
    const jan1 = serialOf(y, 1, 1)
    return Math.floor((s - jan1 + ((weekday(jan1) - first + 7) % 7)) / 7) + 1
  }),
  WORKDAY: atp('e j D?', (start: number, days: number, hol?: number[]) => workday(start, days, weekendOf(1) as boolean[], holidaySet(hol))),
  'WORKDAY.INTL': atp('e j a? D?', (start: number, days: number, w: V | undefined, hol?: number[]) => {
    const weekend = weekendOf(w)
    return isErr(weekend) ? weekend : workday(start, days, weekend, holidaySet(hol))
  }),
  YEAR: { args: 'd', fn: (s: number) => ymd(s).y },
  YEARFRAC: atp('e e j?', (start: number, end: number, basis = 0) => (validBasis(basis) ? yearFrac(start, end, basis) : E.NUM)),
}
