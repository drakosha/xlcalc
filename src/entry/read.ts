/**
 * A constant typed into a cell, as Excel reads it in a book's locale (`Rules`), measured in the entry set `typed`: a
 * number with the format its form asks for (`1,234.5` → `#,##0.00`, `5%` → `0%`, `$5` → the locale's money with red
 * negatives, `1.5E+3` → scientific, `1 1/2` → a fraction), a date or a time with the built-in date and time formats, a
 * boolean by the locale's words in any case, an error by the locale's text in any case, else text as typed. Numbers are
 * read as VALUE reads them, dates and times as a date argument does (no space before them); a euro sign before or after
 * a number is dropped in every locale, and a number in scientific form past the range of numbers is text with the
 * scientific format.
 */
import type { NumberFormat } from '../contract.js'
import { literal } from '../criteria.js'
import { dateTimeParts, trimSpaces, type DateTimeParts } from '../dates.js'
import { readsClock, type RulesClock } from '../locale/rules/data.js'
import type { Rules } from '../locale/rules.js'
import { ERROR_LITERALS } from '../locale/lex.js'

/** What a typed number is by its form: a cell's format follows it (`forms.ts`). */
export type Form = 'plain' | 'grouped' | 'percent' | 'currency' | 'scientific' | 'fraction' | 'date' | 'time'

export type Typed =
  | { t: 'number'; v: number; form: Form; format?: NumberFormat }
  | { t: 'boolean'; v: boolean }
  /** An error literal, invariant (`#N/A`). */
  | { t: 'error'; literal: string }
  /** Text; `format` where Excel sets one all the same (a number in scientific form out of range). */
  | { t: 'text'; v: string; format?: NumberFormat }

/** The error literals a typed constant can be: those a formula reads (`#SPILL!` and newer are text). */
const TYPED_ERRORS = ERROR_LITERALS.slice(0, 7)

/** The locale's money format, as Excel sets it for typed money: grouped, two decimals or none, negatives in red. */
export function moneyFormat(c: Rules['currency'], decimals: boolean): string {
  const num = decimals ? '#,##0.00' : '#,##0'
  const sym = `"${c.symbol}"`
  const gap = c.space ? '\\ ' : ''
  const pos = c.before ? `${sym}${gap}${num}` : `${num}${gap}${sym}`
  return c.negative === '()' ? `${pos}_);[Red]\\(${pos}\\)` : `${pos};[Red]\\-${pos}`
}

/** The format of a number's form: money, scientific, percent, a fraction, grouping; two decimals where a decimal separator or a fraction was typed. */
function numberForm(text: string, r: Rules): { form: Form; format?: NumberFormat } {
  const s = trimSpaces(text)
  const frac = /\d \d+\/(\d+)/.exec(s)
  const decimals = frac !== null || new RegExp(`\\d*${literal(r.decimal)}`).test(s.replace(/[eE][+-]?\d+$/, ''))
  if (s.includes(r.currency.symbol)) return { form: 'currency', format: moneyFormat(r.currency, decimals) }
  if (/\d[eE][+-]?\d/.test(s)) return { form: 'scientific', format: 11 }
  if (s.includes('%')) return { form: 'percent', format: decimals ? 10 : 9 }
  if (frac) return { form: 'fraction', format: frac[1].length > 1 ? 13 : 12 }
  if (new RegExp(`\\d${literal(r.group)}\\d`).test(s)) return { form: 'grouped', format: decimals ? 4 : 3 }
  return { form: 'plain' }
}

/** A time's format by what was typed: AM or PM, seconds, a fraction of a second, hours past a day. */
function timeFormat(t: string, decimal: string): NumberFormat {
  const ampm = /[ap]m?$/.test(t)
  const parts = t.replace(/ ?[ap]m?$/, '').split(/ ?: ?/)
  const seconds = parts.length > 2 && parts[2] !== ''
  if (ampm) return seconds ? 19 : 18
  if (t.includes(decimal)) return 47
  if (Number(parts[0]) >= 24) return 46
  return seconds ? 21 : 20
}

/** A date's format by its parts: a day and a month (of the clock's year), a month and a year, a month by name, a time with it. */
function dateFormat(p: DateTimeParts, yearless: boolean, decimal: string): NumberFormat {
  if (p.date === null) return timeFormat(p.time!, decimal)
  if (p.time !== null) return 22
  if (yearless) return 16
  const numbers = p.date.match(/\d+/g)?.length ?? 0
  if (/\p{L}/u.test(p.date)) return numbers >= 2 ? 15 : 17
  return numbers >= 3 ? 14 : 17
}

/** A number of the locale with a euro sign before or after it, inside parentheses too (`€5`, `5 €`, `-€5`, `(€5)`): the number, the sign dropped; not with percent. */
function euro(text: string, r: Rules): number | null {
  if (r.currency.symbol === '€' || text.includes('%') || text.split('€').length !== 2) return null
  return r.numberOnly(text.replace(/ *€ */, ''))
}

/** A fraction without a whole part typed into a percent cell (`1/2`, which its editor makes `1/2%`): that part of a hundred. */
function fractionPercent(text: string): number | null {
  const m = /^ *([+-]?)(\d+)\/(\d+) *$/.exec(text)
  return m && Number(m[3]) > 0 ? ((m[1] === '-' ? -1 : 1) * Number(m[2])) / Number(m[3]) / 100 : null
}

/**
 * Number text in scientific form that Excel keeps as text, in the scientific format: an exponent past 308 or below
 * -324, a magnitude of 1E+308 or more (Excel's largest typed number is below it), a mantissa that is not 0 and comes to
 * 0 (`1E-330`).
 */
function outOfRange(text: string, n: number | null, r: Rules): boolean {
  const m = new RegExp(`^ *[+-]?(\\d*(?:${literal(r.decimal)}\\d*)?)[eE]([+-]?\\d+) *$`).exec(text)
  if (!m || !/\d/.test(m[1])) return false
  const e = Number(m[2])
  return e > 308 || e < -324 || n === null || Math.abs(n) >= 1e308 || (n === 0 && /[1-9]/.test(m[1]))
}

/** The smallest normal number: Excel takes a smaller one typed as 0 (`1E-309`). */
const MIN_NORMAL = 2.2250738585072014e-308

/**
 * Typed text read as a constant by the book's rules; `clock` is the rules' clock, to know a date without a year. In a
 * percent cell (`percent`) a number begun by a digit gets the `%` Excel's editor adds as it is typed.
 */
export function readTyped(typed: string, r: Rules, clock: RulesClock, percent = false): Typed {
  const editor = percent && /^[-+]?\d/.test(typed)
  const fp = editor ? fractionPercent(typed) : null
  if (fp !== null) return { t: 'number', v: fp, form: 'percent', format: 10 }
  const text = editor ? `${typed}%` : typed
  const n = r.numberOnly(text)
  if (outOfRange(text, n, r)) return { t: 'text', v: text, format: 11 }
  if (n !== null) return { t: 'number', v: Math.abs(n) < MIN_NORMAL ? 0 : n, ...numberForm(text, r) }
  // date and time text is read with no space before it, as an argument reads it
  const [p, yearless] = readsClock(clock, () => dateTimeParts(text, r.dates))
  if (p) return { t: 'number', v: p.value, form: p.date === null ? 'time' : 'date', format: dateFormat(p, yearless, r.dates.decimal) }
  const e = euro(text, r)
  if (e !== null) return { t: 'number', v: e, form: 'plain' }
  const u = text.toUpperCase()
  const { TRUE, FALSE } = r.locale.booleans
  if (u === TRUE.toUpperCase() || u === FALSE.toUpperCase()) return { t: 'boolean', v: u === TRUE.toUpperCase() }
  const err = TYPED_ERRORS.find((e) => u === (r.locale.errors[e] ?? e).toUpperCase())
  return err ? { t: 'error', literal: err } : { t: 'text', v: text }
}
