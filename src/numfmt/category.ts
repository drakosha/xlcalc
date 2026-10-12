/**
 * The category of a number format as the Number Format box of Excel's Home tab shows it, and the items of that box's
 * list, in a locale (`numberFormatCategory`, `numberFormatChoices`); measured in Excel 16.113 for Mac
 * (`test/excel/ui/ribbon.json`). The box names the category of the Format Cells window whose formats hold the code, as
 * Excel keeps it (`canonical`: a literal written bare is kept escaped, `$` as `\$`, so `$#,##0.00` is not the Currency
 * `"$"#,##0.00`), else Custom. Excel builds most categories' formats in the locale's own syntax and compares them with the
 * en-US code, so in a locale that writes formats otherwise (ru-RU: `0,00`, `[Красный]`, `ДД.ММ.ГГГГ`) only those its
 * syntax leaves as they are match (`0`, `0%`); bracketed currencies, their accounting formats and the special formats
 * are kept in one table for every locale and match in each.
 */
import type { FormatCategory, FormatChoice, FormatChoiceKey, NumberFormat, NumberFormatCategory } from '../contract.js'
import { checkFormat } from '../input.js'
import { checkLocaleArg } from '../locale/load.js'
import type { FormatNames } from '../locale/rules/data.js'
import type { Locale } from '../locale/types.js'
import { builtinCodes } from './builtin.js'
import { splitSections } from './parse.js'

/** en-US's names of the box's categories and the list's items, and the code of its Currency item. */
const EN_NAMES: FormatNames = {
  categories: { general: 'General', number: 'Number', currency: 'Currency', accounting: 'Accounting', date: 'Date', time: 'Time', percentage: 'Percentage', fraction: 'Fraction', scientific: 'Scientific', text: 'Text', special: 'Special', custom: 'Custom' },
  choices: { general: 'General', number: 'Number', currency: 'Currency', accounting: 'Accounting', shortDate: 'Short Date', longDate: 'Long Date', time: 'Time', percentage: 'Percentage', fraction: 'Fraction', scientific: 'Scientific', text: 'Text' },
  currency: '"$"#,##0.00',
}

/** The system's long date and time, as Excel keeps them (`[$-x-sysdate]` reads as F800). */
const LONG_DATE = '[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy'
const LONG_TIME = '[$-F400]h:mm:ss\\ AM/PM'

/** The list's items in its order, each with what it applies: a built-in format by number, or a code (the Currency item's is the locale's). */
const CHOICES: readonly [FormatChoiceKey, NumberFormat | undefined][] = [
  ['general', 0],
  ['number', 2],
  ['currency', undefined],
  ['accounting', 44],
  ['shortDate', 14],
  ['longDate', LONG_DATE],
  ['time', LONG_TIME],
  ['percentage', 10],
  ['fraction', 12],
  ['scientific', 11],
  ['text', 49],
]

const optLocale = (locale: string | Locale | undefined) => (locale === undefined ? undefined : checkLocaleArg(locale, 'locale'))
const namesOf = (loc: Locale | undefined): FormatNames => loc?.regional?.formats ?? EN_NAMES
const codesOf = (loc: Locale | undefined) => builtinCodes(loc?.regional?.builtin ?? {})

/** The money formats among the built-in ones, whose `$` Excel keeps quoted (`"$"#,##0_);\("$"#,##0\)`), as the box reads them. */
const MONEY = /^(?:[5-8]|4[1-4])$/

/**
 * The category of a number format in `locale` (a loaded locale's tag or a locale object; en-US without one) as the
 * ribbon's Number Format box shows it, with its name there: a code as an xlsx file stores it, or a built-in format by
 * its number.
 */
export function numberFormatCategory(format: NumberFormat, locale?: string | Locale): NumberFormatCategory {
  checkFormat(format, 'format')
  const loc = optLocale(locale)
  const codes = codesOf(loc)
  const code = typeof format === 'string' ? format : MONEY.test(String(format)) ? codes(format).replace(/\$/g, '"$"') : codes(format)
  const category = categoryOf(canonical(code), codes(14), loc)
  return { category, name: namesOf(loc).categories[category] }
}

/** The items of the Number Format list in `locale`, in its order: each one's key, name, the format it applies and that format's code in the locale. */
export function numberFormatChoices(locale?: string | Locale): FormatChoice[] {
  const loc = optLocale(locale)
  const names = namesOf(loc)
  const codes = codesOf(loc)
  return CHOICES.map(([key, format = names.currency]) => ({ key, name: names.choices[key], format, code: typeof format === 'number' ? codes(format) : format }))
}

/** A date or time letter outside quotes and escapes: `y m d h s`. */
const DATE_LETTER = /[ymdhs]/i

/**
 * A code as Excel keeps it: each literal written bare escaped (`-` as `\-`, a space as `\ `, `$` as `\$`; in a date `.`
 * and `,` too), quoted text, escapes and brackets as written, `[$-x-sysdate]` as `[$-F800]`.
 */
export function canonical(code: string): string {
  return splitSections(code.replace(/\[\$-x-sysdate\]/gi, '[$-F800]').replace(/\[\$-x-systime\]/gi, '[$-F400]'))
    .map((s) => {
      const date = DATE_LETTER.test(s.replace(/"[^"]*"|\\.|\[[^\]]*\]|[_*].|General|AM\/PM|A\/P|E[+-]/gi, ''))
      let out = ''
      for (let i = 0; i < s.length; i++) {
        const c = s[i]
        const end = c === '"' ? s.indexOf('"', i + 1) : c === '[' ? s.indexOf(']', i) : c === '\\' || c === '_' || c === '*' ? i + 1 : -1
        const word = /^(?:General|AM\/PM|A\/P|E[+-])/i.exec(s.slice(i))?.[0]
        const to = end >= 0 ? end : word ? i + word.length - 1 : -1
        if (to >= 0) {
          out += s.slice(i, to + 1)
          i = to
        } else out += /[0-9#?%@/:]/.test(c) || DATE_LETTER.test(c) || (!date && (c === '.' || c === ',')) || (c === '.' && s[i + 1] === '0') ? c : `\\${c}`
      }
      return out
    })
    .join(';')
}

/** Digits as Format Cells writes them, 0 to 30 decimals: `0.00`, and with groups `#,##0.00`. */
const NUM = String.raw`(?:0|#,##0)(?:\.0{1,30})?`
/** Money's digits: always grouped. */
const MONEY_NUM = String.raw`#,##0(?:\.0{1,30})?`
/** A currency symbol of the table every locale has: `[$€-407]`, `[$USD]`. */
const BRACKET = String.raw`\[\$[^\]-]+(?:-[^\]]*)?\]`

const NUMBER = new RegExp(String.raw`^(${NUM})(?:;\[Red\]\1|_\);(?:\[Red\])?\\\(\1\\\))?$`, 'i')
const PERCENT = /^0(?:\.0{1,30})?%$/
const SCIENTIFIC = /^0\.0{1,30}E\+00$/
const FRACTION = /^#\\ (?:\?\/\?|\?\?\/\?\?|\?\?\?\/\?\?\?|\?\/[248]|\?\?\/16|\?\/10|\?\?\/100)$/
const SPECIAL: ReadonlySet<string> = new Set(['00000', '00000\\-0000', '000\\-00\\-0000', '000000', '[<=9999999]###\\-####;\\(###\\)\\ ###\\-####'])

/** Currency with the symbol `s` (a regular expression) before the digits, or a bracketed one after them with a space; negatives with a minus in red, or in parentheses. */
const currency = (s: string, after: boolean) => {
  const p = `(?:${s}(?:\\\\ )?${MONEY_NUM}${after ? `|${MONEY_NUM}\\\\ ${s}` : ''})`
  return new RegExp(String.raw`^(${p})(?:;\[Red\]\\-\1|_\);(?:\[Red\])?\\\(\1\\\))?$`, 'i')
}

/** Accounting with the symbol `s` before the fill (or none): `_(S* #,##0.00_);_(S* \(#,##0.00\);_(S* "-"??_);_(@_)`; with a bracketed one, also Excel for Mac's forms with minus signs, the symbol before or after. */
const accounting = (s: string, minus: boolean) => {
  const paren = String.raw`_\((${s})\* (${MONEY_NUM})_\);_\(\1\* \\\(\2\\\);_\(\1\* "-"\?*_\);_\(@_\)`
  if (!minus) return new RegExp(`^${paren}$`)
  const before = String.raw`_-(${s})\* (${MONEY_NUM})_ ;_-\3\* \\-\4\\ ;_-\3\* "-"\?*_ ;_-@_ `
  const after = String.raw`_-\* (${MONEY_NUM})\\ (${s})_-;\\-\* \5\\ \6_-;_-\* "-"\?*\\ \6_-;_-@_-`
  return new RegExp(`^(?:${paren}|${before}|${after})$`)
}

const BRACKET_CURRENCY = currency(BRACKET, true)
const BRACKET_ACCOUNTING = accounting(`${BRACKET}(?:\\\\ )?`, true)
const EN_CURRENCY = currency('"\\$"', false)
const EN_ACCOUNTING = accounting('(?:"\\$")?', false)

/** Date and time as Format Cells writes most of them: one date or time section (a locale tag but the system's), then `;@`. */
const WITH_TEXT = /^(?:\[\$-(?!F[48]00)[0-9A-F]+\])?((?:"[^"]*"|\[[hms]+\]|[^;[\]"])+);@$/i

/** A date or time without numbers' placeholders but fractions of seconds: Date when it has a day, a year or months alone, else Time. */
function dateOrTime(section: string): FormatCategory | undefined {
  const bare = section.replace(/"[^"]*"|\\./g, '')
  if (/[#?%,@]|E[+-]|General/i.test(bare) || /\d/.test(bare.replace(/\.0+/g, '')) || !DATE_LETTER.test(bare)) return undefined
  return /[dy]/i.test(bare) || !/[hs]|AM\/PM|A\/P/i.test(bare) ? 'date' : 'time'
}

/**
 * Whether a code reads the same in a locale's own syntax: Excel builds most categories' formats in it, so a code
 * with a decimal point, a group separator, a space (the group separator of ru-RU), a colour or date letters that the
 * locale writes otherwise matches none of them. en-US and a locale without rules of computation write as xlsx does.
 */
function sameIn(code: string, loc: Locale | undefined): boolean {
  const r = loc?.regional
  if (!r) return true
  const bare = code.replace(/"[^"]*"/g, '"').replace(/\\(.)/g, (_, c: string) => (c === ' ' ? ' ' : '\\'))
  if ((r.decimal !== '.' && bare.includes('.')) || (r.group !== ',' && bare.includes(',')) || (/\s/.test(r.group) && bare.includes(' '))) return false
  if (r.colors[2] !== 'Red' && /\[(?:Black|White|Red|Green|Blue|Yellow|Magenta|Cyan|Color\d+)\]/i.test(bare)) return false
  return !(r.codes.day !== 'd' && DATE_LETTER.test(bare.replace(/\[[^\]]*\]|General|AM\/PM|A\/P|E[+-]/gi, '')))
}

/** The category of a code as Excel keeps it (`canonical`) in a locale whose short date (built-in 14) has the code `short`. */
function categoryOf(code: string, short: string, loc: Locale | undefined): FormatCategory {
  if (/^General$/i.test(code)) return 'general'
  if (code === '@') return 'text'
  if (SPECIAL.has(code)) return 'special'
  if (BRACKET_CURRENCY.test(code)) return 'currency'
  if (BRACKET_ACCOUNTING.test(code)) return 'accounting'
  if (!sameIn(code, loc)) return 'custom'
  if (NUMBER.test(code)) return 'number'
  if (PERCENT.test(code)) return 'percentage'
  if (SCIENTIFIC.test(code)) return 'scientific'
  if (FRACTION.test(code)) return 'fraction'
  if (!loc?.regional && EN_CURRENCY.test(code)) return 'currency'
  if (!loc?.regional && EN_ACCOUNTING.test(code)) return 'accounting'
  // the locale's short date as the package shows built-in 14 (Excel for Windows's), and Excel for Mac's own in en-US
  if (code === canonical(short) || code === 'm/d/yy' || code === LONG_DATE) return 'date'
  if (code === LONG_TIME) return 'time'
  const m = WITH_TEXT.exec(code)
  return (m && dateOrTime(m[1])) || 'custom'
}
