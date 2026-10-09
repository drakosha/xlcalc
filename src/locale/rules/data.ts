/**
 * What a locale other than en-US gives its rules of computation (`regional.ts`), beside its formula text for TRUE,
 * FALSE and error texts (`LocaleFile`): number separators, the currency format, month and day names, TEXT's format
 * codes, the text order, the code page of CHAR and CODE, the word of new table columns' names; and the clock that
 * completes a date written without a year, whose reads make a formula depend on the clock.
 */

export interface RegionalData {
  /** The decimal separator of numbers in text, and the group separator (a plain space in ru-RU). */
  readonly decimal: string
  readonly group: string
  /** DOLLAR's and VALUE's currency: its symbol after the amount, a space between, a negative amount with a minus. */
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
  /** CHAR's and CODE's characters of codes 128–255: the locale's Windows code page, an undefined byte the C1 control of its value. */
  readonly codePage: string
  /** The word of a new table column's name (`Столбец1`), as Excel names a column inserted into a table or a cleared header. */
  readonly columnWord: string
}

/** The clock as the rules read it: date text without a year takes the clock's year, and reading it sets `read`. */
export interface RulesClock {
  year(): number
  read: boolean
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
