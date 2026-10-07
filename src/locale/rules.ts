/**
 * The rules of computation of a book's locale, as Excel's regional settings change values: text read as a
 * number, a date or a boolean, numbers and booleans turned into text, TEXT's format codes, DOLLAR's and
 * FIXED's formats, the order of text, numbers and words in criteria. The engine builds one rules object per
 * book (`rulesOf`) and passes it to the operators and criteria (`runtime.ts`) and to functions (`Cx.rules`);
 * code outside a book (the parser, the criteria index of `unstable/internal`) works by en-US.
 */
import type { CellDisplay, CellValue, ErrorCode } from '../contract.js'
import { EN_CRIT, type CritRules } from '../criteria.js'
import { parseDateText, textToNumber } from '../dates.js'
import { displayWith, EN_CODES, EN_NAMES, formatWith, generalFormat, generalText, type Dialect, type Names, type Style } from '../numfmt.js'
import { NUMBER_TEXT, readNumber } from '../model.js'
import { ERROR_TEXT } from '../syntax.js'
import { compareOp, compareValues, toBool, toNum, toText, type CErr, type V } from '../values.js'
import { enUS } from './en-US.js'
import type { RegionalData, RulesClock } from './rules/data.js'
import { regionalRules } from './rules/regional.js'
import { RU_RU } from './rules/ru-RU.js'
import type { Locale } from './types.js'

/** The data of the locales other than en-US, by tag (one line per locale of `LOCALES`). */
export const REGIONAL: ReadonlyMap<string, RegionalData> = new Map([['ru-RU', RU_RU]])

export interface Rules extends CritRules {
  /** The locale of formula text: the words for TRUE and FALSE, error texts, the separators ARRAYTOTEXT writes. */
  readonly locale: Locale
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
  /** What a cell shows: a value by its format (an xlsx code) in `space` points of room. */
  display(value: CellValue, format: string | undefined, space: number): CellDisplay
  /** Excel's order of two values that are not errors (lookups, sorting), text by the locale's order. */
  compare(a: V, b: V): number
  /** The order of the comparison operators: `compare` with numbers equal to 15 significant digits. */
  compareOp(a: V, b: V): number
  /** An error as text (VALUETOTEXT, ARRAYTOTEXT). */
  errorText(code: ErrorCode): string
  /** A real number in complex text (`1.5` in `1.5+2i`), or null. */
  complexPart(text: string): number | null
}

/** The month and day names of a locale code (`[$-409]`): en-US's, or a regional locale's. */
const namesOf = (lcid: string): Names | undefined => (lcid === '409' ? EN_NAMES : [...REGIONAL.values()].find((d) => d.lcid === lcid))

/** en-US's style: the measured long date and time, `/` for dates in TEXT and in cells alike. */
const EN_STYLE: Style = {
  ...EN_NAMES,
  decimal: '.',
  group: ',',
  dateSeparator: '/',
  general: generalFormat,
  bool: (b) => (b ? 'TRUE' : 'FALSE'),
  error: (code) => ERROR_TEXT[code as ErrorCode],
  names: namesOf,
  longDate: 'dddd, mmmm d, yyyy',
  longTime: 'h:mm:ss AM/PM',
}

const EN_DIALECT: Dialect = {
  codes: EN_CODES,
  style: EN_STYLE,
  number: (text) => {
    const t = text.trim()
    return NUMBER_TEXT.test(t) ? readNumber(t) : textToNumber(text)
  },
}

/** Excel en-US with the English interface: the engine's own rules since the first release. */
export const EN_US: Rules = {
  ...EN_CRIT,
  locale: enUS,
  decimal: '.',
  group: ',',
  currency: { symbol: '$', before: true, space: '', negative: '()' },
  toNum,
  toText,
  toBool,
  valueText: textToNumber,
  dateText: parseDateText,
  numberText: generalText,
  formatText: (value, format) => formatWith(value, format, EN_DIALECT),
  display: (value, format, space) => displayWith(value, format, EN_STYLE, space),
  compare: compareValues,
  compareOp,
  errorText: (code) => ERROR_TEXT[code],
  complexPart: (t) => (NUMBER_TEXT.test(t) ? readNumber(t) : null),
}

/** The rules of each locale for showing values outside a book (the clock is not read there). */
const DISPLAY = new Map<string, Rules>()

/** Rules that show values in the locale `tag` (one of `LOCALES`; en-US without one), kept for later calls. */
export function displayRules(tag: string | undefined): Rules {
  const data = tag === undefined ? undefined : REGIONAL.get(tag)
  if (!data) return EN_US
  let r = DISPLAY.get(tag!)
  if (!r) DISPLAY.set(tag!, (r = regionalRules(data, { year: () => 1900, read: false }, namesOf)))
  return r
}

/** The rules of a book in the locale `tag` (one of `LOCALES`; en-US without one), reading `clock` for dates without a year. */
export function rulesOf(tag: string | undefined, clock: RulesClock): Rules {
  const data = tag === undefined ? undefined : REGIONAL.get(tag)
  return data ? regionalRules(data, clock, namesOf) : EN_US
}
