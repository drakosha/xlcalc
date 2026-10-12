/**
 * What a locale other than en-US gives its rules of computation (`regional.ts`), beside its formula text for TRUE,
 * FALSE and error texts (`Locale.regional`): number separators, the currency format, month and day names, TEXT's format
 * codes, the text order, the code page of CHAR and CODE, the word of new table columns' names; and the clock that
 * completes a date written without a year, whose reads make a formula depend on the clock.
 */
import type { FormatCategory, FormatChoiceKey } from '../../contract.js'

/** A code page of CHAR and CODE: its name as Unicode publishes its table (`windows-1251`), and its characters of codes 128–255, an undefined byte the C1 control of its value. */
export interface CodePage {
  readonly name: string
  readonly chars: string
}

/** The Number Format box of the Home tab in a locale: the names of the categories it shows and of its list's items, and the code its Currency item applies. */
export interface FormatNames {
  readonly categories: Readonly<Record<FormatCategory, string>>
  readonly choices: Readonly<Record<FormatChoiceKey, string>>
  readonly currency: string
}

/** The order of day, month and year in date text (`dates.ts`). */
export type DateOrder = 'dmy' | 'mdy' | 'ymd'

export interface RegionalData {
  /** The decimal separator of numbers in text, and the group separator (a plain space in ru-RU). */
  readonly decimal: string
  readonly group: string
  /** The order of day, month and year in date text of numbers: `dmy` (`31.01.2025`), `mdy` (`1/31/2025`) or `ymd` (`2025/1/31`). */
  readonly dateOrder: DateOrder
  /** DOLLAR's and VALUE's currency: the symbol, before the amount or after it (VALUE reads it there), the space between, a negative amount in parentheses or with a minus. */
  readonly currency: { readonly symbol: string; readonly before: boolean; readonly space: string; readonly negative: '()' | '-' }
  /** Month names in the nominative as TEXT writes them (`MMMM`), short ones (`MMM`); date text reads a prefix of a full name. */
  readonly months: readonly string[]
  readonly monthsShort: readonly string[]
  /** Day names from Sunday as TEXT writes them (`DDDD`, `DDD`). */
  readonly days: readonly string[]
  readonly daysShort: readonly string[]
  /** TEXT's format codes: the letters of day, month, minutes, year, hours, seconds, the name of the General format, the word of `[Color10]`. */
  readonly codes: { readonly day: string; readonly month: string; readonly minute: string; readonly year: string; readonly hour: string; readonly second: string; readonly general: string; readonly color: string }
  /** TEXT's colour names: black, white, red, green, blue, yellow, magenta, cyan. */
  readonly colors: readonly string[]
  /** What a bare `/` in a cell's date format shows (`31.01.2025`). */
  readonly dateSeparator: string
  /** Excel's locale code (`[$-419]`), upper-case hex. */
  readonly lcid: string
  /** The long date and time (`[$-F800]`, `[$-F400]`) as en-US codes, and the long date's month names. */
  readonly longDate: string
  readonly longTime: string
  readonly longMonths: readonly string[]
  /** The `Intl.Collator` tag that orders text as this Excel does. */
  readonly collation: string
  /** The built-in formats the locale shows by codes of its own, by number, as en-US codes (`numfmt/builtin.ts`). */
  readonly builtin: Readonly<Record<number, string>>
  /** CHAR's and CODE's code page: the locale's Windows one. */
  readonly codePage: CodePage
  /** The word of a new table column's name (`Столбец1`), as Excel names a column inserted into a table or a cleared header. */
  readonly columnWord: string
  /** The names of the Number Format box (`numberFormatCategory`, `numberFormatChoices`); en-US's without them. */
  readonly formats?: FormatNames
}

/** The clock as the rules read it: date text without a year takes the clock's year, and reading it sets `read`. */
export interface RulesClock {
  year(): number
  read: boolean
}

/** A clock of the rules whose year is `year()`: asking for it marks the clock read. */
export function markingClock(year: () => number): RulesClock {
  const clock: RulesClock = {
    read: false,
    year: () => {
      clock.read = true
      return year()
    },
  }
  return clock
}

/** `f()`, and whether it read the clock's year; an earlier read stays marked. Without a clock nothing reads it. */
export function readsClock<T>(clock: RulesClock | undefined, f: () => T): [T, boolean] {
  if (!clock) return [f(), false]
  const was = clock.read
  clock.read = false
  const v = f()
  const used = clock.read
  clock.read = was || used
  return [v, used]
}
