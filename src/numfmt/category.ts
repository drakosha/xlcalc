/**
 * The category of a number format as the Number Format box of Excel's Home tab shows it, and the items of that box's
 * list with the codes they apply, in a locale (`numberFormatCategory`, `numberFormatChoices`). A code is read by the
 * engine's parser (`parse.ts`) and matched against the formats each category of Excel's Format Cells window makes; one
 * that no category makes is Custom. en-US's names are built in, another locale's come with it (`RegionalData.formats`).
 */
import type { FormatCategory, FormatChoice, FormatChoiceKey, NumberFormat, NumberFormatCategory } from '../contract.js'
import { checkFormat } from '../input.js'
import { checkLocaleArg } from '../locale/load.js'
import type { FormatNames } from '../locale/rules/data.js'
import type { Locale } from '../locale/types.js'
import { builtinCodes } from './builtin.js'
import { EN_CODES, parseFormat, type Section, type Tok } from './parse.js'

/** en-US's names: the box's categories and the list's items, and its Currency item's code. */
const EN_NAMES: FormatNames = {
  categories: { general: 'General', number: 'Number', currency: 'Currency', accounting: 'Accounting', date: 'Date', time: 'Time', percentage: 'Percentage', fraction: 'Fraction', scientific: 'Scientific', text: 'Text', special: 'Special', custom: 'Custom' },
  choices: { general: 'General', number: 'Number', currency: 'Currency', accounting: 'Accounting', shortDate: 'Short Date', longDate: 'Long Date', time: 'Time', percentage: 'Percentage', fraction: 'Fraction', scientific: 'Scientific', text: 'Text' },
  currency: '$#,##0.00',
}

/** The list's items in its order, each with what it applies: a built-in format by number, or a code. */
const CHOICES: readonly [FormatChoiceKey, FormatCategory, NumberFormat | undefined][] = [
  ['general', 'general', 0],
  ['number', 'number', 2],
  ['currency', 'currency', undefined],
  ['accounting', 'accounting', 44],
  ['shortDate', 'date', 14],
  ['longDate', 'date', '[$-F800]dddd, mmmm dd, yyyy'],
  ['time', 'time', '[$-F400]h:mm:ss AM/PM'],
  ['percentage', 'percentage', 10],
  ['fraction', 'fraction', 12],
  ['scientific', 'scientific', 11],
  ['text', 'text', 49],
]

const namesOf = (loc: Locale | undefined): FormatNames => loc?.regional?.formats ?? EN_NAMES
const codesOf = (loc: Locale | undefined) => builtinCodes(loc?.regional?.builtin ?? {})

const optLocale = (locale: string | Locale | undefined) => (locale === undefined ? undefined : checkLocaleArg(locale, 'locale'))

/** The category of a number format (a code, or a built-in format's number shown by the locale's code) in `locale` (a loaded locale's tag or a locale object; en-US without one), and its name as the box shows it. */
export function numberFormatCategory(format: NumberFormat, locale?: string | Locale): NumberFormatCategory {
  checkFormat(format, 'format')
  const loc = optLocale(locale)
  const code = typeof format === 'number' ? codesOf(loc)(format) : format
  const category = categoryOf(code)
  return { category, name: namesOf(loc).categories[category] }
}

/** The items of the Number Format list in `locale`, in its order: each one's key, its category, its name and the format it applies, with that format's code in the locale. */
export function numberFormatChoices(locale?: string | Locale): FormatChoice[] {
  const loc = optLocale(locale)
  const names = namesOf(loc)
  const code = codesOf(loc)
  return CHOICES.map(([choice, category, format = names.currency]) => ({ choice, category, name: names.choices[choice], format, code: typeof format === 'number' ? code(format) : format }))
}

/** Digits as a category of Format Cells writes them: `0`, `#,##0`, with decimals `.00`; the count of decimals is free. */
const DIGITS = /^(?:0|#,##0)(?:\.0+)?$/

/** A section's tokens as text, literals quoted, for matching. */
function shape(toks: readonly Tok[]): string {
  return toks
    .map((t) => {
      switch (t.t) {
        case 'lit':
          return t.s === ' ' ? ' ' : `"${t.s}"`
        case 'ph':
          return t.c
        case 'dot':
          return '.'
        case 'comma':
          return ','
        case 'pct':
          return '%'
        case 'slash':
          return '/'
        case 'exp':
          return t.plus ? 'E+' : 'E-'
        case 'date':
        case 'elapsed':
          return t.s
        case 'ampm':
          return 'AM/PM'
        case 'general':
          return 'G'
        case 'text':
          return '@'
        case 'pad':
          return `_${t.s}`
        case 'fill':
          return `*${t.s}`
      }
    })
    .join('')
}

const isDate = (t: Tok) => t.t === 'date' || t.t === 'elapsed' || t.t === 'ampm'

/** The category of a code (en-US syntax), as the box shows it. */
export function categoryOf(code: string): FormatCategory {
  const secs = parseFormat(code, EN_CODES)
  if (!secs) return 'custom'
  const [first] = secs
  const s = secs.map((x) => shape(x.toks))
  if (secs.length === 1 && s[0] === 'G' && !first.color && !first.cond) return 'general'
  if (secs.length === 1 && s[0] === '@' && !first.color) return 'text'
  if (secs.some((x) => x.cond)) return 'custom'
  if (first.toks.some(isDate)) return dateKind(secs)
  if (secs.length === 1 && !first.color) {
    if (/^0(?:\.0+)?%$/.test(s[0])) return 'percentage'
    if (/^0(?:\.0+)?E\+00$/.test(s[0])) return 'scientific'
    if (/^# (?:\?{1,3}\/\?{1,3}|\?\/[248]|\?\?\/16|\?\/10|\?\?\/100)$/.test(s[0])) return 'fraction'
    if (/^0{5}(?:"-"0{4})?$/.test(s[0]) || s[0] === '000"-"00"-"0000') return 'special'
  }
  if (numberLike(secs, s)) return 'number'
  return moneyKind(secs) ?? 'custom'
}

/** Number: one section of digits, or a second for negatives with a minus, in red, or in parentheses. */
function numberLike(secs: readonly Section[], s: readonly string[]): boolean {
  if (secs[0].color || secs.length > 2 || !DIGITS.test(s[0].replace(/_\)$/, ''))) return false
  if (secs.length === 1) return !s[0].endsWith('_)')
  const d = s[0].replace(/_\)$/, '')
  const neg = secs[1]
  if (neg.color && neg.color !== '#FF0000') return false
  const paren = s[0].endsWith('_)')
  return paren ? s[1] === `"("${d}")"` : s[1] === d || s[1] === `"-"${d}`
}

/** Currency and Accounting: digits with a currency symbol; Accounting fills the room between them (`* `). */
function moneyKind(secs: readonly Section[]): FormatCategory | undefined {
  const symbol = secs[0].toks.some((t) => t.t === 'lit' && /[$€£¥₽]|[^\x20-\x7e]/.test(t.s) && t.s.trim() !== '')
  if (!symbol || secs.some((x) => x.color && x.color !== '#FF0000')) return undefined
  return secs[0].toks.some((t) => t.t === 'fill') ? 'accounting' : 'currency'
}

/** Date and Time: a date with a time is Date; elapsed time and times alone are Time; months alone a date. */
function dateKind(secs: readonly Section[]): FormatCategory {
  if (secs.length > 2 || secs[0].color) return 'custom'
  const toks = secs[0].toks
  if (toks.some((t) => t.t === 'date' && 'yd'.includes(t.s[0]))) return 'date'
  return toks.some((t) => t.t === 'elapsed' || t.t === 'ampm' || (t.t === 'date' && 'hs'.includes(t.s[0]))) ? 'time' : 'date'
}
