/**
 * The category of a number format as the Number Format box of Excel's Home tab shows it, and the items of that box's
 * list, in a locale (`numberFormatCategory`, `numberFormatChoices`); measured in Excel 16.113 for Mac in en-US
 * (`test/excel/ui/ribbon.json`) and in Excel 16.0 for Windows in ru-RU (`test/excel/win/ribbon/ru-RU/`). The box names
 * the category of the Format Cells window whose formats hold the code, as Excel keeps it (`canonical`: a literal
 * written bare is kept escaped, `$` as `\$`, so `$#,##0.00` is not the Currency `"$"#,##0.00`), else Custom. Excel for
 * Windows writes the code in the locale's syntax and compares it with its formats written so: in ru-RU `0.00` is
 * Числовой, as in en-US, but money is the ruble's quoted after the amount, negatives take a minus, a fraction's space is
 * quoted. Bracketed currencies (`[$€-407]`), their accounting formats and the special formats are every locale's.
 * Excel for Mac in a locale other than en-US compares the code, kept in en-US syntax, with its formats in the locale's
 * (`0,00`), so it shows Custom for nearly every code: the package follows Excel for Windows there.
 */
import type { FormatCategory, FormatChoice, FormatChoiceKey, NumberFormat, NumberFormatCategory } from '../contract.js'
import { checkFormat } from '../input/styles.js'
import { optionalLocale } from '../locale/load.js'
import type { FormatNames } from '../locale/rules/data.js'
import type { Locale } from '../locale/types.js'
import { builtinCodes, EN_BUILTIN } from './builtin.js'
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

const namesOf = (loc: Locale | undefined): FormatNames => loc?.regional?.formats ?? EN_NAMES
const codesOf = (loc: Locale | undefined) => builtinCodes(loc?.regional?.builtin ?? {})

/** The money formats among the built-in ones, whose `$` Excel keeps quoted (`"$"#,##0_);\("$"#,##0\)`), as the box reads them. */
const MONEY = /^(?:[5-8]|4[1-4])$/

/**
 * The category of a number format in `locale` (a loaded locale's tag or a locale object; en-US without one), as the
 * ribbon's Number Format box displays it, with the category's name there: for a format code as stored in an xlsx file,
 * or a built-in format by its number.
 */
export function numberFormatCategory(format: NumberFormat, locale?: string | Locale): NumberFormatCategory {
  checkFormat(format, 'format')
  const loc = optionalLocale(locale, 'locale')
  let category: FormatCategory
  if (!loc?.regional) category = categoryOf(typeof format === 'string' ? canonical(format) : EN_KEPT[format], '', undefined)
  else category = typeof format === 'number' ? builtinIn(format) : categoryOf(canonical(format), codesOf(loc)(14), loc)
  return { category, name: namesOf(loc).categories[category] }
}

/** The items of the Number Format list in `locale`, in list order: each item's key, name, the format it applies and that format's code in the locale. */
export function numberFormatChoices(locale?: string | Locale): FormatChoice[] {
  const loc = optionalLocale(locale, 'locale')
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

const NUMBER = new RegExp(String.raw`^(${NUM})(?:;\[Red\]\1)?$`, 'i')
/** Number's negatives in parentheses, en-US's alone. */
const PAREN = new RegExp(String.raw`^(${NUM})_\);(?:\[Red\])?\\\(\1\\\)$`, 'i')
const PERCENT = /^0(?:\.0{1,30})?%$/
const SCIENTIFIC = /^0\.0{1,30}E\+00$/
const DENOMINATORS = String.raw`(?:\?\/\?|\?\?\/\?\?|\?\?\?\/\?\?\?|\?\/[248]|\?\?\/16|\?\/10|\?\?\/100)`
const FRACTION = new RegExp(String.raw`^#\\ ${DENOMINATORS}$`)
/** Fractions as a locale other than en-US writes them, the space quoted (ru-RU's built-in 12, `#" "?/?`): one escaped is no fraction there. */
const QUOTED_FRACTION = new RegExp(String.raw`^#" "${DENOMINATORS}$`)
const SPECIAL: ReadonlySet<string> = new Set(['00000', '00000\\-0000', '000\\-00\\-0000', '000000', '[<=9999999]###\\-####;\\(###\\)\\ ###\\-####'])

/** Currency with the symbol `s` (a regular expression) right before the digits, or a bracketed one before them (a space between or none) or after them with a space; negatives with a minus in red, or in parentheses. */
const currency = (s: string, bracketed: boolean) => {
  const p = bracketed ? `(?:${s}(?:\\\\ )?${MONEY_NUM}|${MONEY_NUM}\\\\ ${s})` : `${s}${MONEY_NUM}`
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
/** Negatives with a minus, in red or not, as a locale other than en-US writes them. */
const MINUS = String.raw`;(?:\[Red\])?\\-\1`
const BRACKET_MINUS = new RegExp(String.raw`^(${BRACKET}(?:\\ )?${MONEY_NUM}|${MONEY_NUM}\\ ${BRACKET})${MINUS}$`, 'i')

const LOCAL_MONEY = new Map<string, { currency: RegExp; accounting: RegExp }>()
/**
 * A locale's own money, its symbol quoted after the digits with a space (ru-RU's built-in 5 to 8 and 41 to 44 and its
 * Currency item, `#,##0.00\ "₽"`): currency with negatives with a minus, accounting with the symbol or none.
 */
function localMoney(symbol: string) {
  let m = LOCAL_MONEY.get(symbol)
  if (!m) {
    const s = String.raw`\\ "${symbol.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`
    m = {
      currency: new RegExp(String.raw`^(${MONEY_NUM}${s})(?:${MINUS})?$`, 'i'),
      accounting: new RegExp(String.raw`^_-\* (${MONEY_NUM})((?:${s})?)_-;\\-\* \1\2_-;_-\* "-"\?*\2_-;_-@_-$`),
    }
    LOCAL_MONEY.set(symbol, m)
  }
  return m
}

/** Date and time as Format Cells writes most of them: one date or time section (a locale tag but the system's), then `;@`. */
const WITH_TEXT = /^(?:\[\$-(?!F[48]00)[0-9A-F]+\])?((?:"[^"]*"|\[[hms]+\]|[^;[\]"])+);@$/i

/** A date or time without numbers' placeholders but fractions of seconds: Date when it has a day, a year or months alone, else Time. */
function dateOrTime(section: string): FormatCategory | undefined {
  const bare = section.replace(/"[^"]*"|\\./g, '')
  if (/[#?%,@]|E[+-]|General/i.test(bare) || /\d/.test(bare.replace(/\.0+/g, '')) || !DATE_LETTER.test(bare)) return undefined
  return /[dy]/i.test(bare) || !/[hs]|AM\/PM|A\/P/i.test(bare) ? 'date' : 'time'
}

/** en-US's code of each built-in format as Excel keeps it, the dollar of its money quoted. */
const EN_KEPT = EN_BUILTIN.map((code, id) => canonical(MONEY.test(String(id)) ? code.replace(/\$/g, '"$"') : code))

/**
 * A built-in format's category in a locale other than en-US, as Excel for Windows shows it by the locale's own code
 * (measured in ru-RU): its money the locale's currency (5 to 8) and accounting (41 to 44), its numbers with a minus
 * (37 to 40) none of Format Cells' formats, the others as their en-US codes.
 */
const builtinIn = (id: number): FormatCategory => (id >= 37 && id <= 40 ? 'custom' : MONEY.test(String(id)) ? (id < 9 ? 'currency' : 'accounting') : categoryOf(EN_KEPT[id], '', undefined))

/**
 * The category of a code as Excel keeps it (`canonical`) in a locale whose short date (built-in 14) has the code
 * `short`. A locale other than en-US (measured in ru-RU) takes en-US's numbers, percents, exponents, text, dates and
 * times with `;@`, and its own money, fractions with the space quoted and short date with `/` for its separator
 * (ru-RU's `dd/mm/yyyy`, shown `ДД.ММ.ГГГГ`); en-US's money, Number's negatives in parentheses and fractions with the
 * space escaped are none of its formats, nor is a built-in format's en-US code written as a code (`m/d/yyyy`).
 */
function categoryOf(code: string, short: string, loc: Locale | undefined): FormatCategory {
  const r = loc?.regional
  if (/^General$/i.test(code)) return 'general'
  if (code === '@') return 'text'
  if (SPECIAL.has(code)) return 'special'
  if (BRACKET_CURRENCY.test(code)) return 'currency'
  if (BRACKET_ACCOUNTING.test(code)) return 'accounting'
  if (NUMBER.test(code) || (!r && PAREN.test(code))) return 'number'
  if (PERCENT.test(code)) return 'percentage'
  if (SCIENTIFIC.test(code)) return 'scientific'
  if (r ? QUOTED_FRACTION.test(code) : FRACTION.test(code)) return 'fraction'
  const local = r && localMoney(r.currency.symbol)
  if (local ? local.currency.test(code) || BRACKET_MINUS.test(code) : EN_CURRENCY.test(code)) return 'currency'
  if (local ? local.accounting.test(code) : EN_ACCOUNTING.test(code)) return 'accounting'
  // the short date: en-US's built-in 14 and Excel for Mac's own, or the locale's with `/`
  if (code === LONG_DATE || (r ? code === short.split(r.dateSeparator).join('/') : code === 'm/d/yyyy' || code === 'm/d/yy')) return 'date'
  if (code === LONG_TIME) return 'time'
  const m = WITH_TEXT.exec(code)
  return (m && dateOrTime(m[1])) || 'custom'
}
