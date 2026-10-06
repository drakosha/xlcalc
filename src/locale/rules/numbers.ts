/**
 * Number text as a locale other than en-US reads it in VALUE, in arithmetic and in numeric arguments, measured
 * in ru-RU: the group separator before groups of three digits or more (`1 234,5`, `1 2345`; not `12 34`), the
 * decimal separator, a sign, parentheses, `%` at either end, an exponent, `1 1/2`, the currency symbol after the
 * amount (`1 234,5 ₽`, not with `%`); else a time or a date by the locale's date syntax.
 */
import { dateTimeWith, timeOf, trimSpaces, unwrap, type DateSyntax } from '../../dates.js'
import { literal } from '../../criteria.js'
import { readNumber } from '../../model.js'
import type { RegionalData } from './data.js'

export function numberReader(data: RegionalData, dates: DateSyntax): (text: string) => number | null {
  const { decimal, group } = data
  const { symbol } = data.currency
  const grouped = new RegExp(`^((?!0+${literal(group)})\\d+(?:${literal(group)}\\d{3,})*|)(${literal(decimal)}\\d*)?$`)
  return (text) => {
    const t = trimSpaces(text)
    if (t === '') return null
    const { s: unwrapped, paren, percent } = unwrap(t)
    const currency = unwrapped.endsWith(symbol)
    if (currency && percent) return null
    const s = currency ? trimSpaces(unwrapped.slice(0, -symbol.length)) : unwrapped
    const sign = /^([+-]?) *(.*)$/.exec(s)!
    const negative = sign[1] === '-'
    if (paren && negative) return null
    const body = sign[2]
    const exp = /^(.*?)([eE][+-]?\d+)$/.exec(body)
    const mant = exp ? exp[1] : body
    let n: number | null = null
    if (grouped.test(mant) && /\d/.test(mant)) n = readNumber(mant.split(group).join('').replace(decimal, '.') + (exp ? exp[2] : ''))
    else if (!paren) {
      const frac = /^(\d+) (\d+)\/(\d+)$/.exec(body)
      if (frac && Number(frac[3]) > 0) n = Number(frac[1]) + Number(frac[2]) / Number(frac[3])
      else if (!sign[1] && !percent && !currency) return timeOf(s, decimal) ?? dateTimeWith(s, dates)
    }
    if (n === null || !Number.isFinite(n)) return null
    if (percent) n /= 100
    return negative || paren ? -n : n
  }
}
