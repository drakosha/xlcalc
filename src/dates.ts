/**
 * The Excel 1900 date system in the functions: serial numbers and calendar dates, with Excel's day 0
 * (1900-01-00) and its phantom 1900-02-29 (serial 60), and date, time and number text as Excel en-US reads it;
 * the times and the frame of date text are shared with the other locales (`DateSyntax`, `locale/rules/`).
 */
import { readNumber } from './model.js'

/** The last serial number Excel accepts: 9999-12-31. */
export const MAX_SERIAL = 2_958_465

const DAY_MS = 86_400_000
const EPOCH_MS = Date.UTC(1899, 11, 30)

/** The serial number of a real calendar date (month and day overflow roll over, as in `Date.UTC`). */
export function serialOf(y: number, m: number, d: number): number {
  const t = new Date(0)
  t.setUTCFullYear(y, m - 1, d)
  const serial = Math.round((t.getTime() - EPOCH_MS) / DAY_MS)
  // before 1900-03-01 Excel is one day behind (the phantom 1900-02-29)
  return serial < 61 ? serial - 1 : serial
}

export const isLeap = (y: number) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0

export const daysInMonth = (y: number, m: number) => [31, isLeap(y) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]

/** Year, month and day of a serial number in Excel's calendar: 0 is 1900-01-00, 60 is 1900-02-29. */
/** The day of a date-time serial as DAY, MONTH, YEAR and WEEKDAY read it: to the nearest second first (0.4 s before midnight is the next day). */
export const dayAt = (t: number): number => Math.floor(Math.round(t * 86_400) / 86_400)

export function ymd(serial: number): { y: number; m: number; d: number } {
  const s = Math.floor(serial)
  if (s === 0) return { y: 1900, m: 1, d: 0 }
  if (s === 60) return { y: 1900, m: 2, d: 29 }
  const t = new Date(EPOCH_MS + (s < 61 ? s + 1 : s) * DAY_MS)
  return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() }
}

/** A month count from year 0 normalized to a year and a month (1–12). */
export function addMonths(y: number, m: number, months: number): { y: number; m: number } {
  const total = y * 12 + (m - 1) + months
  const ny = Math.floor(total / 12)
  return { y: ny, m: total - ny * 12 + 1 }
}

// ---------- date and time text (en-US) ----------

const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
const MONTH_RE = '(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sept?(?:ember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)'

/** How a locale reads date text: the date part (lower case, single spaces) and the decimal separator of seconds. */
export interface DateSyntax {
  date(s: string): number | null
  decimal: string
}

/** The patterns of time text by the decimal separator of seconds: a time in date text, date text and its time, VALUE's time, VALUE's date and time. */
interface TimePatterns {
  time: RegExp
  split: RegExp
  value: RegExp
  dateTime: RegExp
}

const PATTERNS = new Map<string, TimePatterns>()

/** The patterns for a decimal separator, built once: `18:30`, `6:30:15 PM`, `6 PM`, seconds with decimals. */
function patterns(decimal: string): TimePatterns {
  let p = PATTERNS.get(decimal)
  if (!p) {
    const d = `\\${decimal}`
    const t = `(\\d{1,4}):(\\d+)(?::(\\d+(?:${d}\\d+)?))?(?:\\s*([ap]m))?|(\\d{1,2})\\s*([ap]m)`
    p = {
      time: new RegExp(`^(?:${t})$`),
      split: new RegExp(`^(.*?)(?:[ t](${t}))?$`),
      value: new RegExp(`^(\\d+):(\\d+)(?::(\\d+(?:${d}\\d+)?))?(?: ([ap]m))?$`, 'i'),
      dateTime: new RegExp(`^(.*\\S) +(\\d+:\\d+(?::\\d+(?:${d}\\d+)?)?(?: [ap]m)?)$`, 'i'),
    }
    PATTERNS.set(decimal, p)
  }
  return p
}
const DATE_FORMS: [RegExp, (m: RegExpExecArray) => [string, string, string]][] = [
  [/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/, (m) => [m[1], m[2], m[3]]],
  [/^(\d{1,2})\/(\d{1,2})\/(\d{4}|\d{2})$/, (m) => [m[3], m[1], m[2]]],
  [new RegExp(`^(\\d{1,2})[- ]${MONTH_RE}[- ](\\d{4}|\\d{2})$`), (m) => [m[3], m[2], m[1]]],
  [new RegExp(`^${MONTH_RE} (\\d{1,2}),? (\\d{4})$`), (m) => [m[3], m[1], m[2]]],
]

/** A year as typed: two digits are 1930–2029, as in Excel. */
export const fullYear = (s: string) => (s.length > 2 ? Number(s) : Number(s) + (Number(s) < 30 ? 2000 : 1900))

/** The serial number of a calendar date from 1900 to 9999, the phantom 1900-02-29 too; null if there is no such date. */
export function calendarDate(y: number, mo: number, d: number): number | null {
  if (y < 1900 || y > 9999 || mo < 1 || mo > 12 || d < 1) return null
  if (y === 1900 && mo === 2 && d === 29) return 60
  return d > daysInMonth(y, mo) ? null : serialOf(y, mo, d)
}

function dateText(s: string): number | null {
  for (const [re, parts] of DATE_FORMS) {
    const m = re.exec(s)
    if (!m) continue
    const [ys, ms, ds] = parts(m)
    return calendarDate(fullYear(ys), /^\d+$/.test(ms) ? Number(ms) : MONTHS.indexOf(ms.slice(0, 3)) + 1, Number(ds))
  }
  return null
}

const EN_DATES: DateSyntax = { date: dateText, decimal: '.' }

function timeText(s: string, decimal: string): number | null {
  const m = patterns(decimal).time.exec(s)
  if (!m) return null
  const ampm = m[4] ?? m[6]
  let h = Number(m[1] ?? m[5])
  const mi = Number(m[2] ?? 0)
  const sec = Number((m[3] ?? '0').replace(decimal, '.'))
  // Excel en-US: with AM/PM hours 0–12 and minutes and seconds below 60; without, up to 9999 hours and any minutes and seconds
  if (ampm) {
    if (h > 12 || mi > 59 || sec >= 60) return null
    h = (h % 12) + (ampm === 'pm' ? 12 : 0)
  }
  return (h * 3600 + mi * 60 + sec) / 86_400
}

/**
 * Date or time text as Excel en-US reads it in a date or time argument: `2025-01-31`, `2025/01/31`,
 * `1/31/2025`, `1/31/25`, `31-Jan-2025`, `Jan 31, 2025`, each optionally followed by a time, or a time
 * alone (`18:30`, `6:30:15 PM`, `6 PM`). The serial number with a day fraction, or null. Forms without a
 * year (`1/31`, which Excel completes with the current year) are not read.
 */
export const parseDateText = (text: string): number | null => parseDateWith(text, EN_DATES)

/** Date or time text in a date or time argument by a locale's syntax: a date optionally followed by a time, or a time alone. */
export function parseDateWith(text: string, syn: DateSyntax): number | null {
  const s = text.trim().toLowerCase().replace(/\s+/g, ' ')
  if (s === '') return null
  const time = timeText(s, syn.decimal)
  if (time !== null) return time
  const split = patterns(syn.decimal).split.exec(s)!
  const date = syn.date(split[1])
  if (date === null) return null
  if (split[2] === undefined) return date
  const t = timeText(split[2], syn.decimal)
  return t === null ? null : date + t
}

// ---------- number text (en-US): VALUE's reading, also of text in arithmetic and numeric arguments ----------

/** Digits with commas, each followed by three digits at least (Excel: `1,0000` is 10000, `10,00` and `0,123` are not numbers). */
const GROUPED = /^((?!0+,)\d+(?:,\d{3,})*|)(\.\d*)?$/

/** Excel trims spaces only, not tabs or line breaks. */
export const trimSpaces = (s: string) => s.replace(/^ +| +$/g, '')

const MAX_TIME = 10_000 / 24

/** VALUE's time, seconds with the decimal separator: minutes and seconds past 59 carry (`1:61`), up to 9,999 hours; with ` AM`/` PM` a clock time. */
export function timeOf(s: string, decimal: string): number | null {
  const m = patterns(decimal).value.exec(s)
  if (!m) return null
  let h = Number(m[1])
  const mi = Number(m[2])
  const sec = Number((m[3] ?? '0').replace(decimal, '.'))
  if (m[4]) {
    if (h > 12 || mi > 59 || sec >= 60) return null
    h = (h % 12) + (m[4].toLowerCase() === 'pm' ? 12 : 0)
  }
  const t = (h * 3600 + mi * 60 + sec) / 86_400
  return t < MAX_TIME ? t : null
}

/** VALUE's date text by a locale's syntax, optionally followed by a time that may run past 24 hours (`2025-01-31 25:00`). */
export function dateTimeWith(s: string, syn: DateSyntax): number | null {
  const m = patterns(syn.decimal).dateTime.exec(s)
  if (m) {
    const d = m[1].includes(':') ? null : parseDateWith(m[1], syn)
    const t = timeOf(m[2], syn.decimal)
    return d === null || t === null ? null : d + t
  }
  return s.includes(':') || /\d[ap]m$/i.test(s) ? null : parseDateWith(s, syn)
}

/**
 * Text as a number, as Excel en-US reads it in VALUE, in arithmetic and in numeric arguments: sign, `$`,
 * parentheses, commas, exponent, one `%` at either end, `1 1/2`; else a time or a date. Null if it is none.
 */
export function textToNumber(text: string): number | null {
  let s = trimSpaces(text)
  if (s === '') return null
  let paren = false
  let percent = false
  // parentheses and the percent sign in either order: `(5%)`, `(5)%`, `%(5)`
  for (;;) {
    if (!paren && s.startsWith('(') && s.endsWith(')')) [paren, s] = [true, s.slice(1, -1)]
    else if (!percent && s.startsWith('%')) [percent, s] = [true, s.slice(1)]
    else if (!percent && s.endsWith('%')) [percent, s] = [true, s.slice(0, -1)]
    else break
    s = trimSpaces(s)
  }
  const sign = /^([+-]?) *(\$?) *([+-]?) *(.*)$/.exec(s)!
  if (sign[1] && sign[3]) return null
  const negative = (sign[1] || sign[3]) === '-'
  if (paren && negative) return null
  const body = sign[4]
  const exp = /^(.*?)([eE][+-]?\d+)$/.exec(body)
  const mant = exp ? exp[1] : body
  let n: number | null = null
  if (GROUPED.test(mant) && /\d/.test(mant)) n = readNumber(mant.replace(/,/g, '') + (exp ? exp[2] : ''))
  else if (!sign[2] && !paren) {
    const frac = /^(\d+) (\d+)\/(\d+)$/.exec(body)
    if (frac && Number(frac[3]) > 0) n = Number(frac[1]) + Number(frac[2]) / Number(frac[3])
    else if (!sign[1] && !sign[3] && !percent) return timeOf(s, '.') ?? dateTimeWith(s, EN_DATES)
  }
  if (n === null || !Number.isFinite(n)) return null
  if (percent) n /= 100
  return negative || paren ? -n : n
}
