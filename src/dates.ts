/**
 * The Excel 1900 date system in the functions: serial numbers and calendar dates, with Excel's day 0
 * (1900-01-00) and its phantom 1900-02-29 (serial 60), and date and time text as Excel en-US reads it.
 */

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
const TIME_RE = '(\\d{1,4}):(\\d+)(?::(\\d+(?:\\.\\d+)?))?(?:\\s*([ap]m))?|(\\d{1,2})\\s*([ap]m)'
const DATE_FORMS: [RegExp, (m: RegExpExecArray) => [string, string, string]][] = [
  [/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/, (m) => [m[1], m[2], m[3]]],
  [/^(\d{1,2})\/(\d{1,2})\/(\d{4}|\d{2})$/, (m) => [m[3], m[1], m[2]]],
  [new RegExp(`^(\\d{1,2})[- ]${MONTH_RE}[- ](\\d{4}|\\d{2})$`), (m) => [m[3], m[2], m[1]]],
  [new RegExp(`^${MONTH_RE} (\\d{1,2}),? (\\d{4})$`), (m) => [m[3], m[1], m[2]]],
]

/** A year as typed: two digits are 1930–2029, as in Excel. */
const fullYear = (s: string) => (s.length > 2 ? Number(s) : Number(s) + (Number(s) < 30 ? 2000 : 1900))

function dateText(s: string): number | null {
  for (const [re, parts] of DATE_FORMS) {
    const m = re.exec(s)
    if (!m) continue
    const [ys, ms, ds] = parts(m)
    const y = fullYear(ys)
    const mo = /^\d+$/.test(ms) ? Number(ms) : MONTHS.indexOf(ms.slice(0, 3)) + 1
    const d = Number(ds)
    if (y < 1900 || y > 9999 || mo < 1 || mo > 12 || d < 1) return null
    if (y === 1900 && mo === 2 && d === 29) return 60
    return d > daysInMonth(y, mo) ? null : serialOf(y, mo, d)
  }
  return null
}

function timeText(s: string): number | null {
  const m = new RegExp(`^(?:${TIME_RE})$`).exec(s)
  if (!m) return null
  const ampm = m[4] ?? m[6]
  let h = Number(m[1] ?? m[5])
  const mi = Number(m[2] ?? 0)
  const sec = Number(m[3] ?? 0)
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
export function parseDateText(text: string): number | null {
  const s = text.trim().toLowerCase().replace(/\s+/g, ' ')
  if (s === '') return null
  const time = timeText(s)
  if (time !== null) return time
  const split = new RegExp(`^(.*?)(?:[ t](${TIME_RE}))?$`).exec(s)!
  const date = dateText(split[1])
  if (date === null) return null
  if (split[2] === undefined) return date
  const t = timeText(split[2])
  return t === null ? null : date + t
}
