/**
 * What a locale other than en-US gives its rules of computation (`regional.ts`): the formula-text locale for
 * TRUE, FALSE and error texts, number separators, the currency format, month and day names, TEXT's format
 * codes, the text order; and the clock that completes a date written without a year, whose reads make a
 * formula depend on the clock.
 */
import type { Locale } from '../types.js'

export interface RegionalData {
  /** The formula-text locale: the words for TRUE and FALSE, error texts, the separators ARRAYTOTEXT writes. */
  readonly locale: Locale
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
  /** TEXT's format codes: the letters of day, month, minutes, year, hours, seconds, and the name of the General format. */
  readonly codes: { readonly day: string; readonly month: string; readonly minute: string; readonly year: string; readonly hour: string; readonly second: string; readonly general: string }
  /** The `Intl.Collator` tag that orders text as this Excel does. */
  readonly collation: string
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
