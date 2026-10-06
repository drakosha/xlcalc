/**
 * The rules of computation of a locale other than en-US, built from its data (`RegionalData`, ru-RU's in
 * `ru-RU.ts`) and the book's clock: the same coercions as en-US with the locale's number and date text, words,
 * error texts, separators, currency, TEXT codes and text order.
 */
import type { ErrorCode } from '../../contract.js'
import { errorWords, textOrder } from '../../criteria.js'
import { parseDateWith, type DateSyntax } from '../../dates.js'
import { formatWith, generalText } from '../../format.js'
import { ERROR_TEXT } from '../../syntax.js'
import { compareWith, E, isErr, operatorOrder, type CErr, type V } from '../../values.js'
import type { Rules } from '../rules.js'
import type { RegionalData, RulesClock } from './data.js'
import { dmyDates } from './dates.js'
import { formatDialect } from './format.js'
import { numberReader } from './numbers.js'

export function regionalRules(data: RegionalData, clock: RulesClock): Rules {
  const { locale, decimal, group, currency } = data
  const dates: DateSyntax = { date: dmyDates(data, clock), decimal }
  const valueText = numberReader(data, dates)
  const numberText = (n: number) => generalText(n).replace('.', decimal)
  const { TRUE, FALSE } = locale.booleans
  const word = (t: string): boolean | null => {
    const u = t.toUpperCase()
    return u === TRUE.toUpperCase() ? true : u === FALSE.toUpperCase() ? false : null
  }
  const compareText = textOrder(data.collation)
  const compare = compareWith(compareText)
  const dialect = formatDialect(data, valueText, numberText)
  return {
    locale,
    decimal,
    group,
    currency,
    clock,
    toNum: (v: V): number | CErr => (typeof v === 'number' ? v : v === null ? 0 : typeof v === 'boolean' ? (v ? 1 : 0) : isErr(v) ? v : (valueText(v) ?? E.VALUE)),
    toText: (v: V): string | CErr => (typeof v === 'string' ? v : v === null ? '' : typeof v === 'boolean' ? (v ? TRUE : FALSE) : isErr(v) ? v : numberText(v)),
    toBool: (v: V): boolean | CErr => (typeof v === 'boolean' ? v : typeof v === 'number' ? v !== 0 : v === null ? false : isErr(v) ? v : (word(v) ?? E.VALUE)),
    valueText,
    dateText: (text) => parseDateWith(text, dates),
    numberText,
    formatText: (value, format) => formatWith(value, format, dialect),
    compare,
    compareOp: operatorOrder(compare),
    errorText: (code: ErrorCode) => locale.errors[ERROR_TEXT[code]] ?? ERROR_TEXT[code],
    complexPart: valueText,
    critNumber: valueText,
    critBool: word,
    cellNumber: (s) => (/\d/.test(s) ? valueText(s) : null),
    critError: errorWords(locale.errors),
    compareText,
  }
}
