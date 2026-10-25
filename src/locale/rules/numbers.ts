/**
 * Number text as a locale other than en-US reads it in VALUE, in arithmetic and in numeric arguments, measured
 * in ru-RU: the group separator before groups of three digits or more (`1 234,5`, `1 2345`; not `12 34`), the
 * decimal separator, a sign, parentheses, `%` at either end, an exponent, `1 1/2`, the currency symbol after the
 * amount (`1 234,5 ₽`, not with `%`), spaces around it. Dates and times are read apart (`textReaders` in `dates.ts`).
 * A locale that writes its symbol before the amount (`currency.before`) has it read there, a sign before or after it
 * (`-€5`, `€-5`), as en-US reads `$`; no Excel of such a locale has been measured.
 */
import { bodyNumber, trimSpaces, unwrap } from '../../dates.js'
import { literal } from '../../wildcards.js'
import type { RegionalData } from './data.js'

/** A cell's text as a criterion compares it with a number, by `number`: only text with a digit in it is one. */
export const cellNumberOf = (number: (text: string) => number | null) => (text: string) => (/\d/.test(text) ? number(text) : null)

// `s`: text with a line break is matched whole and read as no number, not left unmatched
const SIGN = /^([+-]?) *(.*)$/s

export function numberReader(data: RegionalData): (text: string) => number | null {
  const { decimal, group } = data
  const { symbol } = data.currency
  const grouped = new RegExp(`^((?!0+${literal(group)})\\d+(?:${literal(group)}\\d{3,})*|)(${literal(decimal)}\\d*)?$`)
  // the symbol before the amount, one sign before or after it
  const before = data.currency.before ? new RegExp(`^([+-]?) *(${literal(symbol)})? *([+-]?) *(.*)$`, 's') : null
  return (text) => {
    const t = trimSpaces(text)
    if (t === '') return null
    const { s, paren, percent } = unwrap(t)
    let negative: boolean
    let body: string
    let currency: boolean
    if (before) {
      const m = before.exec(s)!
      if (m[1] && m[3]) return null
      ;[negative, body, currency] = [(m[1] || m[3]) === '-', m[4], m[2] !== undefined]
    } else {
      currency = s.endsWith(symbol)
      const sign = SIGN.exec(currency ? trimSpaces(s.slice(0, -symbol.length)) : s)!
      ;[negative, body] = [sign[1] === '-', sign[2]]
    }
    if ((currency && percent) || (paren && negative)) return null
    return bodyNumber(body, grouped, (mant) => mant.split(group).join('').replace(decimal, '.'), percent, !paren && !(before && currency), negative || paren)
  }
}
