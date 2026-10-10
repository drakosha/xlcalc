/**
 * A constant typed into a cell, as Excel reads it in a book's locale (`Rules`): a number, with the format its form asks
 * for (`1,234.5` → `#,##0.00`, `5%` → `0%`, `$5` → money, `1.5E+3` → scientific, `1 1/2` → a fraction), a date or a
 * time with the locale's built-in date and time formats, a boolean by the locale's words in any case, an error by its
 * text in the locale or in English, else text as typed. Numbers, dates and times are read as VALUE reads them.
 */
import type { NumberFormat } from '../contract.js'
import { literal } from '../criteria.js'
import { dateTimeParts, trimSpaces, type DateTimeParts } from '../dates.js'
import type { RulesClock } from '../locale/rules/data.js'
import { readsClock } from '../locale/rules/data.js'
import type { Rules } from '../locale/rules.js'
import { ERROR_LITERALS } from '../locale/lex.js'

/** What a typed number is by its form: a cell's format follows it (`forms.ts`). */
export type Form = 'plain' | 'grouped' | 'percent' | 'currency' | 'scientific' | 'fraction' | 'date' | 'time'

export type Typed =
  | { t: 'number'; v: number; form: Form; format?: NumberFormat }
  | { t: 'boolean'; v: boolean }
  /** An error literal, invariant (`#N/A`). */
  | { t: 'error'; literal: string }
  | { t: 'text'; v: string }

/** The error literals a typed constant can be: those a formula reads (`#SPILL!` and newer are text). */
const TYPED_ERRORS = ERROR_LITERALS.slice(0, 7)

/** The format of a number's form, as a built-in format: grouping, percent, money, scientific, fractions by their decimals and denominators. */
function numberForm(text: string, r: Rules): { form: Form; format?: NumberFormat } {
  const s = trimSpaces(text)
  const decimals = new RegExp(`\\d*${literal(r.decimal)}(\\d+)`).exec(s.replace(/[eE][+-]?\d+$/, ''))?.[1].length ?? 0
  if (s.includes(r.currency.symbol)) return { form: 'currency', format: decimals ? 7 : 5 }
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

/** Typed text read as a constant by the book's rules; `clock` is the rules' clock, to know a date without a year. */
export function readTyped(text: string, r: Rules, clock: RulesClock): Typed {
  const n = r.numberOnly(text)
  if (n !== null) return { t: 'number', v: n, ...numberForm(text, r) }
  const [p, yearless] = readsClock(clock, () => dateTimeParts(trimSpaces(text), r.dates))
  if (p) return { t: 'number', v: p.value, form: p.date === null ? 'time' : 'date', format: dateFormat(p, yearless, r.dates.decimal) }
  const u = text.toUpperCase()
  const { TRUE, FALSE } = r.locale.booleans
  if (u === TRUE.toUpperCase() || u === FALSE.toUpperCase()) return { t: 'boolean', v: u === TRUE.toUpperCase() }
  const err = TYPED_ERRORS.find((e) => u === e || u === r.locale.errors[e]?.toUpperCase())
  return err ? { t: 'error', literal: err } : { t: 'text', v: text }
}
