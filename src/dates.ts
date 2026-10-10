/**
 * The Excel 1900 date system in the functions: serial numbers and calendar dates, with Excel's day 0
 * (1900-01-00) and its phantom 1900-02-29 (serial 60), and date, time and number text as Excel en-US reads it;
 * the times and the frame of date text are shared with the other locales (`DateSyntax`, `locale/rules/`).
 */
import { NUMBER_TEXT, readNumber } from './model.js'

/** The last serial number Excel accepts: 9999-12-31. */
export const MAX_SERIAL = 2_958_465

/** The days from the 1900 date system's day 0 to the 1904 system's (1904-01-01): what a date's serial number loses in a 1904 book. */
export const EPOCH_1904 = 1462

/** The day 0 of a book's date system in the 1900 system: 0, or 1462 for 1904. */
export const epochOf = (system: 1900 | 1904 | undefined) => (system === 1904 ? EPOCH_1904 : 0)

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

const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']
/** A month's name, or three letters or more of its beginning (`jan`, `sept`, `febr`). */
const MONTH_RE = '([a-z]{3,})'

/** How a locale reads date text: the date part (lower case, single spaces) and the decimal separator of seconds. */
export interface DateSyntax {
  date(s: string): number | null
  decimal: string
}

/** Date text read in a date system whose day 0 is `epoch` of the 1900 one: a date before it is no date. */
export const inSystem = (syn: DateSyntax, epoch: number): DateSyntax =>
  epoch
    ? {
        ...syn,
        date: (s) => {
          const d = syn.date(s)
          return d === null || d < epoch ? null : d - epoch
        },
      }
    : syn

/** A time's parts as typed: hours, minutes, seconds (with the decimal separator), `a` or `p`. */
type TimeParts = [h: string, mi?: string, sec?: string, ampm?: string]

/** The forms of time text by the decimal separator of seconds (text in lower case, single spaces), and their union for a date with a time. */
interface TimePatterns {
  forms: [RegExp, (m: RegExpExecArray) => TimeParts][]
  dateTime: RegExp
  timeDate: RegExp
}

const PATTERNS = new Map<string, TimePatterns>()

/**
 * The time forms for a decimal separator, built once (the `time-text` set): a clock time `6:30 pm`, `6 :30:15 p`;
 * an hour `6 pm`; a duration `18:30`, `25:00`, `18:30:15.5`, a `:` at the end (`6:`, `6:30:15:`); minutes and
 * seconds with decimals `6:30.5`. Hours of 3 or 4 digits do not begin with 0 (`006:30` is no time).
 */
function patterns(decimal: string): TimePatterns {
  let p = PATTERNS.get(decimal)
  if (!p) {
    const d = `\\${decimal}`
    const h = '(\\d{1,2}|[1-9]\\d{2,3})'
    const sources: [string, (m: RegExpExecArray) => TimeParts][] = [
      [`${h} ?: ?(\\d+)(?: ?: ?(\\d+(?:${d}\\d+)?))? ([ap])m?`, (m) => [m[1], m[2], m[3], m[4]]],
      ['(\\d{1,2}) ([ap])m?', (m) => [m[1], undefined, undefined, m[2]]],
      [`${h} ?:(?: ?(\\d+)(?: ?:(?: ?(\\d+(?:${d}\\d+)?) ?:?)?)?)?`, (m) => [m[1], m[2], m[3]]],
      [`(\\d+) ?: ?(\\d+${d}\\d+)`, (m) => ['0', m[1], m[2]]],
    ]
    const any = sources.map(([src]) => src.replace(/\((?!\?)/g, '(?:')).join('|')
    p = { forms: sources.map(([src, parts]) => [new RegExp(`^${src}$`), parts]), dateTime: new RegExp(`^(.*\\S) (${any})$`), timeDate: new RegExp(`^(${any}) (.*\\S)$`) }
    PATTERNS.set(decimal, p)
  }
  return p
}

/** Full dates; spaces may stand around `/` and `-` (`1 / 31 / 2025`, `1-31-2025`); a year of one digit is of the 2000s (`1-2-3`, the entry set `typed`). */
const DATE_FORMS: [RegExp, (m: RegExpExecArray) => [string, string, string]][] = [
  [/^(\d{4}) *[-/] *(\d{1,2}) *[-/] *(\d{1,2})$/, (m) => [m[1], m[2], m[3]]],
  [/^(\d{1,2}) *[-/] *(\d{1,2}) *[-/] *(\d{4}|\d{1,2})$/, (m) => [m[3], m[1], m[2]]],
  [new RegExp(`^(\\d{1,2})(?: *- *| )${MONTH_RE}(?: *- *| )(\\d{4}|\\d{2})$`), (m) => [m[3], m[2], m[1]]],
  [new RegExp(`^${MONTH_RE} (\\d{1,2}), (\\d{4})$`), (m) => [m[3], m[1], m[2]]],
]

/** A month and a number, no year: `1/31`, `1 / 31`, `1-31`, `Jan 31`, `Jan/31`, `Jan31`, `31-Jan`, `1/2025`, `Jan 2025`. */
const TWO_NUMBERS = /^(\d{1,2}) *[-/] *(\d{1,4})$/
const NAME_NUMBER = new RegExp(`^${MONTH_RE} *[-/]? *(\\d{1,4})$`)
const NUMBER_NAME = new RegExp(`^(\\d{1,2}) *[-/]? *${MONTH_RE}$`)

/** The month a word names (1–12), or 0. */
const monthOf = (word: string) => MONTHS.findIndex((m) => m.startsWith(word)) + 1

/** A year as typed: one or two digits are 1930–2029, as in Excel. */
export const fullYear = (s: string) => (s.length > 2 ? Number(s) : Number(s) + (Number(s) < 30 ? 2000 : 1900))

/** The serial number of a calendar date from 1900 to 9999, the phantom 1900-02-29 too; null if there is no such date. */
export function calendarDate(y: number, mo: number, d: number): number | null {
  if (y < 1900 || y > 9999 || mo < 1 || mo > 12 || d < 1) return null
  if (y === 1900 && mo === 2 && d === 29) return 60
  return d > daysInMonth(y, mo) ? null : serialOf(y, mo, d)
}

/** Whether `d` is a day of month `mo` in some year (February 29 in a leap year). */
const isDay = (mo: number, d: number) => mo >= 1 && mo <= 12 && d >= 1 && d <= daysInMonth(2000, mo)

/**
 * A month and then a number, no year, as Excel en-US reads them (the `date-noyear` set): a day of the month in the
 * current year, `year()`, else a year (two digits 1930–2029, or four) on the 1st, so `2/29` is February 2029 in a
 * year that is not leap and `1/32` January 1932; three digits are neither. A day without `year` is no date.
 */
export function monthAnd(mo: number, n: string, year?: () => number): number | null {
  if (mo < 1 || mo > 12 || n.length === 3) return null
  const day = n.length <= 2 && isDay(mo, Number(n))
  if (day && !year) return null
  return (day ? calendarDate(year!(), mo, Number(n)) : null) ?? calendarDate(fullYear(n), mo, 1)
}

function dateText(s: string, year?: () => number): number | null {
  for (const [re, parts] of DATE_FORMS) {
    const m = re.exec(s)
    if (!m) continue
    const [ys, ms, ds] = parts(m)
    return calendarDate(fullYear(ys), /^\d+$/.test(ms) ? Number(ms) : monthOf(ms), Number(ds))
  }
  let m = TWO_NUMBERS.exec(s)
  if (m) return monthAnd(Number(m[1]), m[2], year)
  if ((m = NAME_NUMBER.exec(s))) return monthAnd(monthOf(m[1]), m[2], year)
  // a number before the name is a day only (`29 Feb` is no date in a year that is not leap)
  if ((m = NUMBER_NAME.exec(s))) {
    const [mo, d] = [monthOf(m[2]), Number(m[1])]
    return year && isDay(mo, d) ? calendarDate(year(), mo, d) : null
  }
  return null
}

/** en-US date text, completing a date without a year with `year()` (the book's clock; without one such text is no date). */
export const enDates = (year?: () => number): DateSyntax => ({ date: (s) => dateText(s, year), decimal: '.' })

const EN_DATES = enDates()

/** Excel's limit of time text: 10,000 hours. */
const MAX_TIME = 10_000 / 24

/**
 * A time as Excel en-US reads it, in VALUE and in date and time arguments alike (the `time-text` set; `patterns`):
 * with AM or PM hours to 12 and minutes and seconds below 60, without any minutes and seconds below 10,000 hours.
 * The day fraction, past 1 from 24 hours, or null.
 */
function timeText(s: string, decimal: string): number | null {
  for (const [re, parts] of patterns(decimal).forms) {
    const m = re.exec(s)
    if (!m) continue
    const [hs, ms = '0', ss = '0', ampm] = parts(m)
    let h = Number(hs)
    const mi = Number(ms)
    const sec = Number(ss.replace(decimal, '.'))
    if (ampm) {
      if (h > 12 || mi > 59 || sec >= 60) return null
      h = (h % 12) + (ampm === 'p' ? 12 : 0)
    }
    const t = (h * 3600 + mi * 60 + sec) / 86_400
    return t < MAX_TIME ? t : null
  }
  return null
}

/**
 * Date or time text as Excel en-US reads it in a date or time argument, outside a book: `2025-01-31`, `2025/01/31`,
 * `1/31/2025`, `1-31-2025`, `1/31/25`, `31-Jan-2025`, `Jan 31, 2025`, with a time before or after it, or a time
 * alone (`18:30`, `6:30:15 PM`, `6 PM`). The serial number with a day fraction, or null. Forms without a year
 * (`1/31`) need a book's clock (`enDates`) and are not read here.
 */
export const parseDateText = (text: string): number | null => parseDateWith(text, EN_DATES)

/**
 * Date or time text by a locale's syntax, as date and time arguments, operators and numeric arguments read it: a time
 * alone, a date alone, or a date and a time in either order (`1/31/2025 6 PM`, `18:00 1/31/2025`; the time may
 * pass 24 hours). Spaces after it are dropped and runs of spaces inside are one; text that begins with a space is
 * no date or time (`DAY(" 1/31/2025")` is #VALUE!), and a tab is no space, as in Excel (the `lead-spaces` set).
 */
export const parseDateWith = (text: string, syn: DateSyntax): number | null => dateTimeParts(text, syn)?.value ?? null

/** Date or time text read by `parseDateWith`: its value, and its date part and time part as read (lower case, single spaces), null where it has none. */
export interface DateTimeParts {
  value: number
  date: string | null
  time: string | null
}

/** `parseDateWith`'s reading with the parts of the text it read, for a cell's format by what was typed (`entry/`). */
export function dateTimeParts(text: string, syn: DateSyntax): DateTimeParts | null {
  const s = text.replace(/ +$/, '').toLowerCase().replace(/  +/g, ' ')
  if (s === '' || s.startsWith(' ')) return null
  const time = timeText(s, syn.decimal)
  if (time !== null) return { value: time, date: null, time: s }
  const date = syn.date(s)
  if (date !== null) return { value: date, date: s, time: null }
  const p = patterns(syn.decimal)
  const both = (d: string, t: string): DateTimeParts | null => {
    const [dd, tt] = [syn.date(d), timeText(t, syn.decimal)]
    return dd === null || tt === null ? null : { value: dd + tt, date: d, time: t }
  }
  let m = p.dateTime.exec(s)
  const after = m ? both(m[1], m[2]) : null
  if (after !== null) return after
  m = p.timeDate.exec(s)
  return m ? both(m[2], m[1]) : null
}

/** How a locale reads text as a number: VALUE's reading and that of operators, numeric arguments, criteria and TEXT. */
export interface TextReaders {
  value(text: string): number | null
  argument(text: string): number | null
}

/**
 * The readers from a locale's number text (`number`: numbers only) and date syntax. Both take numbers with spaces
 * around them (` 5 `); VALUE takes date and time text after dropping spaces before it, the others do not
 * (`VALUE(" 6:30")` is 6:30, `" 6:30"+0` is #VALUE!), as Excel does in en-US and ru-RU (the `lead-spaces` set).
 */
export const textReaders = (number: (text: string) => number | null, dates: DateSyntax): TextReaders => ({
  value: (text) => number(text) ?? parseDateWith(trimSpaces(text), dates),
  argument: (text) => number(text) ?? parseDateWith(text, dates),
})

// ---------- number text (en-US): VALUE's reading, also of text in arithmetic and numeric arguments ----------

/** Digits with commas, each followed by three digits at least (Excel: `1,0000` is 10000, `10,00` and `0,123` are not numbers). */
const GROUPED = /^((?!0+,)\d+(?:,\d{3,})*|)(\.\d*)?$/

/** Excel trims spaces only, not tabs or line breaks. */
export const trimSpaces = (s: string) => s.replace(/^ +| +$/g, '')

/** Parentheses and the percent sign around number text (trimmed), in either order: `(5%)`, `(5)%`, `%(5)`. */
export function unwrap(text: string): { s: string; paren: boolean; percent: boolean } {
  let s = text
  let paren = false
  let percent = false
  for (;;) {
    if (!paren && s.startsWith('(') && s.endsWith(')')) [paren, s] = [true, s.slice(1, -1)]
    else if (!percent && s.startsWith('%')) [percent, s] = [true, s.slice(1)]
    else if (!percent && s.endsWith('%')) [percent, s] = [true, s.slice(0, -1)]
    else break
    s = trimSpaces(s)
  }
  return { s, paren, percent }
}

/**
 * Number text as Excel en-US reads it in VALUE, in arithmetic and in numeric arguments: sign, `$`, parentheses,
 * commas, exponent, one `%` at either end, `1 1/2`, spaces around it. Null if it is none (dates and times are read apart).
 */
export function enNumber(text: string): number | null {
  const t = trimSpaces(text)
  if (t === '') return null
  const { s, paren, percent } = unwrap(t)
  // `s`: text with a line break is matched whole and read as no number, not left unmatched
  const sign = /^([+-]?) *(\$?) *([+-]?) *(.*)$/s.exec(s)!
  if (sign[1] && sign[3]) return null
  const negative = (sign[1] || sign[3]) === '-'
  if (paren && negative) return null
  const body = sign[4]
  const exp = /^(.*?)([eE][+-]?\d+)$/.exec(body)
  const mant = exp ? exp[1] : body
  let n: number | null = null
  // percent moves the decimal point of the text (`1.1%` is 0.011); a fraction is divided
  if (GROUPED.test(mant) && /\d/.test(mant)) n = readNumber(mant.replace(/,/g, '') + (exp ? exp[2] : ''), percent ? -2 : 0)
  else if (!sign[2] && !paren) {
    const frac = /^(\d+) (\d+)\/(\d+)$/.exec(body)
    if (frac && Number(frac[3]) > 0) n = (Number(frac[1]) + Number(frac[2]) / Number(frac[3])) / (percent ? 100 : 1)
  }
  if (n === null || !Number.isFinite(n)) return null
  return negative || paren ? -n : n
}

/** en-US text as a number outside a book, as operators and arguments read it: date text without a year is no date. */
export const textToNumber = textReaders(enNumber, EN_DATES).argument

/** How operators and numeric arguments read text: plain decimal text, whitespace around it trimmed (the common case), else as `read` reads it. */
export const decimalFirst =
  (read: (text: string) => number | null) =>
  (text: string): number | null => {
    const t = text.trim()
    return NUMBER_TEXT.test(t) ? readNumber(t) : read(text)
  }
