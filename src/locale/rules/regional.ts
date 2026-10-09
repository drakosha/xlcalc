/**
 * The rules of computation of a locale other than en-US, built from its data (`RegionalData`, ru-RU's in
 * `ru-RU.ts`) and the book's clock: the same coercions as en-US with the locale's number and date text, words,
 * error texts, separators, currency, TEXT codes, text order and code page.
 */
import type { ErrorCode } from '../../contract.js'
import { textOrder } from '../../collate.js'
import { errorWords } from '../../criteria.js'
import { inSystem, parseDateWith, textReaders } from '../../dates.js'
import { displayWith, formatWith, generalText, styleIn, type Names } from '../../numfmt.js'
import { ERROR_TEXT } from '../../syntax.js'
import { coercions, compareWith, operatorOrder } from '../../values.js'
import type { Rules } from '../rules.js'
import type { RegionalData, RulesClock } from './data.js'
import { dmyDates } from './dates.js'
import { formatDialects } from './format.js'
import { numberReader } from './numbers.js'

export function regionalRules(data: RegionalData, clock: RulesClock, names: (lcid: string) => Names | undefined, epoch = 0): Rules {
  const { locale, decimal, group, currency } = data
  const dates = inSystem({ date: dmyDates(data, clock), decimal }, epoch)
  const { value: valueText, argument } = textReaders(numberReader(data), dates)
  const numberText = (n: number) => generalText(n).replace('.', decimal)
  const { toNum, toText, toBool } = coercions(argument, locale.booleans, numberText)
  const word = (t: string): boolean | null => {
    const b = toBool(t)
    return typeof b === 'boolean' ? b : null
  }
  const compareText = textOrder(data.collation)
  const compare = compareWith(compareText)
  const formats = formatDialects(data, argument, names)
  const [text, cell] = [{ ...formats.text, style: styleIn(formats.text.style, epoch) }, styleIn(formats.cell, epoch)]
  return {
    epoch,
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
    formatText: (value, format) => formatWith(value, format, text),
    display: (value, format, space) => displayWith(value, format, cell, space),
    compare,
    compareOp: operatorOrder(compare),
    errorText: (code: ErrorCode) => locale.errors[ERROR_TEXT[code]] ?? ERROR_TEXT[code],
    codePage: data.codePage,
    critNumber: argument,
    critBool: word,
    cellNumber: (s) => (/\d/.test(s) ? argument(s) : null),
    critError: errorWords(locale.errors),
    compareText,
  }
}
