/**
 * The rules of computation of a book's locale, as Excel's regional settings change values: text read as a
 * number, a date or a boolean, numbers and booleans turned into text, TEXT's format codes, DOLLAR's and
 * FIXED's formats, the order of text, numbers and words in criteria. The engine builds one rules object per
 * book (`rulesOf`) and passes it to the operators and criteria (`runtime.ts`) and to functions (`Cx.rules`);
 * code outside a book (the parser, the criteria index of `unstable/internal`) works by en-US.
 */
import type { CellDisplay, CellValue, ErrorCode, NumberFormat } from '../contract.js'
import { EN_CRIT, type CritRules } from '../criteria.js'
import { decimalFirst, enCritNumber, enDates, enNumber, enTypedNumber, inSystem, parseDateWith, textReaders, type DateSyntax } from '../dates.js'
import { builtinCodes, displayWith, EN_BUILTIN, EN_CODES, EN_NAMES, fitPoints, formatWith, generalFormat, generalText, styleIn, typedPoints, type Dialect, type Names, type Style } from '../numfmt.js'
import { ERROR_TEXT } from '../syntax.js'
import { coercions, compareOp, compareValues, type CErr, type V } from '../values.js'
import { EN_US } from '../locales/en-US.js'
import { RU_NAMES } from './lcid.js'
import { localeOf } from './registry.js'
import type { CodePage, RulesClock } from './rules/data.js'
import { cellNumberOf } from './rules/numbers.js'
import { regionalRules } from './rules/regional.js'
import type { Locale } from './types.js'

/** The codes a locale (a loaded locale's tag or an object; en-US without one) shows the built-in formats by, by number: its own, en-US's for the rest (`numfmt/builtin.ts`). */
export function builtinIn(where: string | Locale | undefined): (id: number) => string {
  const own = localeOf(where).regional?.builtin
  return own ? builtinCodes(own) : (id) => EN_BUILTIN[id]
}

/** A number format as a code in a locale (`ViewportCell.numberFormat`): a built-in format by the locale's code (`builtinIn`), a code as it is. */
export function formatCodeIn(where: string | Locale | undefined): (format: NumberFormat) => string {
  const builtin = builtinIn(where)
  return (format) => (typeof format === 'number' ? builtin(format) : format)
}

export interface Rules extends CritRules {
  /** The locale of formula text: the words for TRUE and FALSE, error texts, the separators ARRAYTOTEXT writes. */
  readonly locale: Locale
  /** The book's date system as the 1900 serial of its day 0: 0, or 1462 in the 1904 system (date text, TEXT and cells count from it). */
  readonly epoch: number
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
  /** Text as VALUE reads it (a number, percent, currency, a date or a time; spaces before a date dropped), or null. */
  valueText(text: string): number | null
  /** Date or time text as a date or time argument reads it (DAY, EDATE…; DATEVALUE after dropping spaces before it), or null. */
  dateText(text: string): number | null
  /** Number text alone as VALUE reads it before dates (a sign, groups, percent, currency, an exponent, `1 1/2`), or null. */
  numberOnly(text: string): number | null
  /** How the locale reads the date and time text of `valueText` (`dateTimeParts`): a typed entry's format follows its form. */
  readonly dates: DateSyntax
  /** A number as text, as `&` writes it. */
  numberText(n: number): string
  /** TEXT(value, format); null for a format the locale does not read. */
  formatText(value: number | string | boolean, format: string): string | null
  /** What a cell shows: a value by its format (an xlsx code or a built-in format's number) in `space` points of room. */
  display(value: CellValue, format: NumberFormat | undefined, space: number): CellDisplay
  /** The points of the column AutoFit gives a cell holding a value by its format (`fitPoints`); null when it shows nothing. */
  fitPoints(value: CellValue, format: NumberFormat | undefined): number | null
  /** The points Excel widens a column of `now` points to as a value is typed into it (`typedPoints`); null when it keeps them. */
  typedPoints(value: CellValue, format: NumberFormat | undefined, now: number): number | null
  /** Excel's order of two values that are not errors (lookups, sorting), text by the locale's order. */
  compare(a: V, b: V): number
  /** The order of the comparison operators: `compare` with numbers equal to 15 significant digits. */
  compareOp(a: V, b: V): number
  /** An error as text (VALUETOTEXT, ARRAYTOTEXT). */
  errorText(code: ErrorCode): string
  /** CHAR's and CODE's code page: the locale's Windows one. */
  readonly codePage: CodePage
}

/** en-US's code page of CHAR and CODE: Windows-1252, an undefined byte the C1 control of its value (`test/codepage.test.ts`). */
const WINDOWS_1252: CodePage = {
  name: 'windows-1252',
  chars: String.fromCharCode(
    0x20ac, 0x0081, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021, 0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008d, 0x017d, 0x008f,
    0x0090, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x009d, 0x017e, 0x0178,
    ...Array.from({ length: 96 }, (_, i) => 0xa0 + i),
  ),
}

/** The month and day names of a locale code (`[$-409]`, `[$-419]`), the same in a book of any locale (`lcid.ts`). */
const namesOf = (lcid: string): Names | undefined => (lcid === '409' ? EN_NAMES : lcid === '419' ? RU_NAMES : undefined)

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
  builtin: (id) => EN_BUILTIN[id],
}

/** Excel en-US with the English interface: the engine's own rules, date text without a year in the clock's year (`1/31`). */
function englishRules(clock: RulesClock, epoch: number): Rules {
  const dates = inSystem(enDates(() => clock.year()), epoch)
  const { value: valueText, argument } = textReaders(enNumber, dates)
  const crit = textReaders(enCritNumber, dates).argument
  const number = decimalFirst(argument)
  const style = styleIn(EN_STYLE, epoch)
  const dialect: Dialect = { codes: EN_CODES, style, number }
  return {
    epoch,
    ...EN_CRIT,
    ...coercions(number, EN_US.booleans, generalText),
    clock,
    critNumber: crit,
    cellNumber: cellNumberOf(argument),
    locale: EN_US,
    decimal: '.',
    group: ',',
    currency: { symbol: '$', before: true, space: '', negative: '()' },
    valueText,
    dateText: (text) => parseDateWith(text, dates),
    numberOnly: enTypedNumber,
    dates,
    numberText: generalText,
    formatText: (value, format) => formatWith(value, format, dialect),
    display: (value, format, space) => displayWith(value, format, style, space),
    fitPoints: (value, format) => fitPoints(value, format, style),
    typedPoints: (value, format, now) => typedPoints(value, format, style, now),
    compare: compareValues,
    compareOp,
    errorText: (code) => ERROR_TEXT[code],
    codePage: WINDOWS_1252,
  }
}

/** The rules of each locale object and date system for showing values outside a book (the clock is not read there). */
const DISPLAY = new WeakMap<Locale, Map<number, Rules>>()

/** Rules that show values in a locale (a loaded locale's tag or an object; en-US without one) and a date system, kept for later calls. */
export function displayRules(where: string | Locale | undefined, epoch = 0): Rules {
  const locale = localeOf(where)
  let byEpoch = DISPLAY.get(locale)
  if (!byEpoch) DISPLAY.set(locale, (byEpoch = new Map()))
  let r = byEpoch.get(epoch)
  if (!r) byEpoch.set(epoch, (r = rulesOf(locale, { year: () => 1900, read: false }, epoch)))
  return r
}

/**
 * The rules of a book in a locale (a loaded locale's tag or an object; en-US without one, and for a locale without
 * rules of its own) and the date system whose day 0 is `epoch` of the 1900 one, reading `clock` for dates without a year.
 */
export function rulesOf(where: string | Locale | undefined, clock: RulesClock, epoch = 0): Rules {
  const locale = localeOf(where)
  return locale.regional ? regionalRules(locale, locale.regional, clock, namesOf, epoch) : englishRules(clock, epoch)
}
