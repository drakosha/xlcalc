/**
 * The rules of computation of a locale other than en-US, built from its data (`RegionalData`, ru-RU's in
 * `ru-RU.ts`) and the book's clock: the same coercions as en-US with the locale's number and date text, words,
 * error texts, separators, currency, TEXT codes and text order.
 */
import type { ErrorCode } from '../../contract.js'
import { errorWords, textOrder } from '../../criteria.js'
import { parseDateWith, textReaders, type DateSyntax } from '../../dates.js'
import { formatWith, generalText } from '../../format.js'
import { ERROR_TEXT } from '../../syntax.js'
import { coercions, compareWith, operatorOrder } from '../../values.js'
import type { Rules } from '../rules.js'
import type { RegionalData, RulesClock } from './data.js'
import { dmyDates } from './dates.js'
import { formatDialect } from './format.js'
import { numberReader } from './numbers.js'

export function regionalRules(data: RegionalData, clock: RulesClock): Rules {
  const { locale, decimal, group, currency } = data
  const dates: DateSyntax = { date: dmyDates(data, clock), decimal }
  const { value: valueText, argument } = textReaders(numberReader(data), dates)
  const numberText = (n: number) => generalText(n).replace('.', decimal)
  const { toNum, toText, toBool } = coercions(argument, locale.booleans, numberText)
  const word = (t: string): boolean | null => {
    const b = toBool(t)
    return typeof b === 'boolean' ? b : null
  }
  const compareText = textOrder(data.collation)
  const compare = compareWith(compareText)
  const dialect = formatDialect(data, argument, numberText)
  return {
    locale,
    decimal,
    group,
    currency,
    clock,
    toNum,
    toText,
    toBool,
    valueText,
    dateText: (text) => parseDateWith(text, dates),
    numberText,
    formatText: (value, format) => formatWith(value, format, dialect),
    compare,
    compareOp: operatorOrder(compare),
    errorText: (code: ErrorCode) => locale.errors[ERROR_TEXT[code]] ?? ERROR_TEXT[code],
    critNumber: argument,
    critBool: word,
    cellNumber: (s) => (/\d/.test(s) ? argument(s) : null),
    critError: errorWords(locale.errors),
    compareText,
  }
}
