/**
 * The rules of computation of a book's locale, as Excel's regional settings change values: text read as a
 * number, a date or a boolean, numbers and booleans turned into text, TEXT's format codes, DOLLAR's and
 * FIXED's formats, the order of text, numbers and words in criteria. The engine builds one rules object per
 * book (`rulesOf`) and passes it to the operators and criteria (`runtime.ts`) and to functions (`Cx.rules`);
 * code outside a book (the parser, the criteria index of `unstable/internal`) works by en-US.
 */
import { EN_CRIT, type CritRules } from '../criteria.js'
import { parseDateText, textToNumber } from '../dates.js'
import { formatText, generalText } from '../format.js'
import { compareOp, compareValues, toBool, toNum, toText, type CErr, type V } from '../values.js'
import { DEFAULT_LOCALE } from './registry.js'

export interface Rules extends CritRules {
  /** The locale's tag. */
  readonly tag: string
  /** The decimal and group separators of numbers in text (FIXED, DOLLAR, NUMBERVALUE's defaults). */
  readonly decimal: string
  readonly group: string
  /** DOLLAR's currency format: the symbol before or after the amount, the space between, a negative amount in parentheses or with a minus. */
  readonly currency: { readonly symbol: string; readonly before: boolean; readonly space: string; readonly negative: '()' | '-' }
  /** A value as a number: operators and numeric arguments (text as VALUE reads it). */
  toNum(v: V): number | CErr
  /** A value as text: `&` and text arguments. */
  toText(v: V): string | CErr
  /** A value as a boolean: the locale's words for TRUE and FALSE in text. */
  toBool(v: V): boolean | CErr
  /** Text as VALUE reads it (a number, percent, currency, a date or a time), or null. */
  valueText(text: string): number | null
  /** Date or time text as a date or time argument reads it (DATEVALUE, DAY, EDATE…), or null. */
  dateText(text: string): number | null
  /** A number as text, as `&` writes it. */
  numberText(n: number): string
  /** TEXT(value, format); null for a format the locale does not read. */
  formatText(value: number | string | boolean, format: string): string | null
  /** Excel's order of two values that are not errors (lookups, sorting), text by the locale's order. */
  compare(a: V, b: V): number
  /** The order of the comparison operators: `compare` with numbers equal to 15 significant digits. */
  compareOp(a: V, b: V): number
}

/** Excel en-US with the English interface: the engine's own rules since the first release. */
export const EN_US: Rules = {
  ...EN_CRIT,
  tag: DEFAULT_LOCALE,
  decimal: '.',
  group: ',',
  currency: { symbol: '$', before: true, space: '', negative: '()' },
  toNum,
  toText,
  toBool,
  valueText: textToNumber,
  dateText: parseDateText,
  numberText: generalText,
  formatText,
  compare: compareValues,
  compareOp,
}

/** The rules of a book in the locale `tag` (one of `LOCALES`). */
export const rulesOf = (_tag: string | undefined): Rules => EN_US
