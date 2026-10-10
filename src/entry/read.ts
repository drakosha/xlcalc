/**
 * A constant typed into a cell, as Excel reads it in a book's locale (`Rules`), measured in the entry set `typed`: a
 * number with the format its form asks for (`1,234.5` → `#,##0.00`, `5%` → `0%`, `$5` → the locale's money with red
 * negatives, `1.5E+3` → scientific, `1 1/2` → a fraction), a date or a time with the built-in date and time formats, a
 * boolean by the locale's words in any case, an error by the locale's text in any case, else text as typed. Numbers,
 * dates and times are read as VALUE reads them; a euro sign before or after a number is dropped in every locale, and a
 * number in scientific form past the range of numbers is text with the scientific format.
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

/** The format of a number's form: grouping, percent, money, scientific, fractions by their decimals and denominators. */
function numberForm(text: string, r: Rules): { form: Form; format?: NumberFormat } {
  const s = trimSpaces(text)
  const decimals = new RegExp(`\\d*${literal(r.decimal)}(\\d+)`).exec(s.replace(/[eE][+-]?\d+$/, ''))?.[1].length ?? 0
  if (s.includes(r.currency.symbol)) return { form: 'currency', format: moneyFormat(r.currency, decimals > 0) }
  if (s.includes('%')) return { form: 'percent', format: decimals ? 10 : 9 }
  if (/\d[eE][+-]?\d/.test(s)) return { form: 'scientific', format: 11 }
  const frac = /\d \d+\/(\d+)/.exec(s)
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

/** A number of the locale with a euro sign before or after it (`€5`, `5 €`, `-€5`): the number, the sign dropped. */
function euro(text: string, r: Rules): number | null {
  const m = /^ *([+-]?) *€ *(.*?) *$/s.exec(text) ?? /^ *(.*?) *€ *$/s.exec(text)
  if (!m || r.currency.symbol === '€') return null
  const n = r.numberOnly(m.length === 3 ? m[1] + m[2] : m[1])
  return n
}

/** A mantissa and an exponent the numbers do not reach (`1E+309`, `1E-400`): Excel keeps the text, in the scientific format. */
const outOfRange = (text: string, r: Rules) => new RegExp(`^ *[+-]?\\d*(${literal(r.decimal)}\\d*)?[eE][+-]?\\d+ *$`).test(text) && /[1-9]/.test(text.replace(/[eE].*$/, ''))

/** Typed text read as a constant by the book's rules; `clock` is the rules' clock, to know a date without a year. */
export function readTyped(text: string, r: Rules, clock: RulesClock): Typed {
  const n = r.numberOnly(text)
  if (n !== null && (n !== 0 || !outOfRange(text, r))) return { t: 'number', v: n, ...numberForm(text, r) }
  if (outOfRange(text, r)) return { t: 'text', v: text, format: 11 }
  const [p, yearless] = readsClock(clock, () => dateTimeParts(trimSpaces(text), r.dates))
  if (p) return { t: 'number', v: p.value, form: p.date === null ? 'time' : 'date', format: dateFormat(p, yearless, r.dates.decimal) }
  const e = euro(text, r)
  if (e !== null) return { t: 'number', v: e, form: 'plain' }
  const u = text.toUpperCase()
  const { TRUE, FALSE } = r.locale.booleans
  if (u === TRUE.toUpperCase() || u === FALSE.toUpperCase()) return { t: 'boolean', v: u === TRUE.toUpperCase() }
  const err = TYPED_ERRORS.find((e) => u === (r.locale.errors[e] ?? e).toUpperCase())
  return err ? { t: 'error', literal: err } : { t: 'text', v: text }
}
