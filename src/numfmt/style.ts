/** How formatted text writes what a code does not spell out: separators, names, General, booleans; and the marks of fills. */
import type { Tok } from './parse.js'

/** Month and day names. */
export interface Names {
  months: readonly string[]
  monthsShort: readonly string[]
  days: readonly string[]
  daysShort: readonly string[]
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

/** en-US names, the names of format codes' own locale. */
export const EN_NAMES: Names = { months: MONTHS, monthsShort: MONTHS.map((m) => m.slice(0, 3)), days: DAYS, daysShort: DAYS.map((d) => d.slice(0, 3)) }

export interface Style extends Names {
  decimal: string
  group: string
  /** What a bare `/` in a date writes: the locale's date separator in a cell (`31.01.2025` in ru-RU), `/` in TEXT. */
  dateSeparator: string
  /** A number in TEXT's General format (11 characters at most), with the locale's decimal separator. */
  general(n: number): string
  bool(b: boolean): string
  /** An error's text, by its code (`DIV0`). */
  error(code: string): string
  /** Names of another locale (`[$-409]`, `[$-419]`), by its upper-case hex code. */
  names(lcid: string): Names | undefined
  /** The locale's long date and time (`[$-F800]`, `[$-F400]`) as en-US codes, and the long date's month names (ru-RU's genitive). */
  longDate: string
  longTime: string
  longMonths?: readonly string[]
  /** The book's date system as the 1900 serial of its day 0: 1462 in the 1904 system, which shows negative dates and times with a minus. */
  epoch?: number
}

/** A style in a date system whose day 0 is `epoch` of the 1900 one. */
export const styleIn = (style: Style, epoch: number): Style => (epoch ? { ...style, epoch } : style)

/** Where a fill `*x` goes in formatted text, followed by `x`: TEXT drops both, a cell repeats `x` across its room. */
export const FILL = '\uFFFF'

/** A token outside a number's digits as text: a literal, `%`, a padding space, a fill's mark; a separator as the locale writes it. */
export function plain(t: Tok, style: Style): string {
  switch (t.t) {
    case 'lit':
      return t.s
    case 'pct':
      return '%'
    case 'slash':
      return '/'
    case 'pad':
      return ' '
    case 'fill':
      return FILL + t.s
    case 'ph':
      return t.c === '0' ? '0' : ''
    case 'dot':
      return style.decimal
    case 'comma':
      return style.group
    default:
      return ''
  }
}
