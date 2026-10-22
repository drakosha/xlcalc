/**
 * Number formats in a locale other than en-US, measured in ru-RU. TEXT reads the locale's codes: the decimal
 * separator (`0,00`), the group separator for grouping and scaling (`# ##0`, `0 ` thousands; a space before the
 * currency symbol is a literal, `# ##0 ₽`), the locale's date and time letters (`ДД.ММ.ГГГГ ч:мм:сс`: day, year,
 * hours and seconds in either case, `М` the month or, after hours or before seconds, minutes as en-US `m`, `м`
 * minutes always), the General format by its name, among literals too (`Основной "шт."`), its colour names
 * (`[Красный]`, `[Цвет10]`), `AM/PM`. Latin letters are literals. Excel refuses a point before a digit placeholder
 * in a number format, where the point is no separator of the locale, and the group separator (a space) in a fraction
 * format. A cell reads an xlsx file's en-US codes and writes the
 * locale's separators, names and date separator (`dd/mm/yyyy` shows `31.01.2025`).
 */
import type { ErrorCode } from '../../contract.js'
import { colorNames } from '../../numfmt/palette.js'
import { builtinCodes, generalFormat, type Codes, type Dialect, type Names, type Style } from '../../numfmt.js'
import { errorIn } from '../../syntax/errors.js'
import type { Locale } from '../types.js'
import type { RegionalData } from './data.js'

/** A section without its quoted text and escaped characters. */
const bare = (sec: string) => sec.replace(/"[^"]*"?|\\[\s\S]/g, '')

/** The locale's codes, as TEXT reads them. */
function regionalCodes(data: RegionalData): Codes {
  const { codes, decimal, group } = data
  // letters read in either case; the month and minute letters differ by case only
  const letters = new Map([codes.day, codes.year, codes.hour, codes.second].map((c, i) => [c.toLowerCase(), 'dyhs'[i]]))
  const letter = (ch: string) => (ch === codes.month ? 'm' : ch === codes.minute ? 'n' : letters.get(ch.toLowerCase()))
  const pointRefused = decimal !== '.' && group !== '.'
  return {
    decimal,
    group,
    general: codes.general.toLowerCase(),
    colors: colorNames(data.colors),
    colorWord: codes.color.toLowerCase(),
    letter,
    same: (a, b) => (a === codes.month || a === codes.minute ? a === b : a.toLowerCase() === b.toLowerCase()),
    currency: data.currency.symbol,
    // a point before a digit placeholder in a number section, where the point is no separator of the locale (ru-RU), or
    // the group separator in a fraction section
    refuses: (sec) => {
      const b = bare(sec)
      // a date by its letters outside brackets, or by elapsed time in them (`[ч]`); a colour's name holds letters too (`[Красный]`)
      const isLetter = (ch: string) => letter(ch) !== undefined
      const elapsed = [...b.matchAll(/\[([^\]]+)\]/g)].some((m) => [...m[1]].every(isLetter))
      const date = elapsed || [...b.replace(/\[[^\]]*\]?/g, '')].some(isLetter) || /AM\/PM|A\/P/i.test(b)
      return (!date && pointRefused && /\.[^]*[0#?]/.test(b)) || (/[0#?]\//.test(b) && b.includes(group))
    },
  }
}

/** TEXT's dialect of a locale, and the style a cell's display writes with (its date separator the locale's). */
export function formatDialects(locale: Locale, data: RegionalData, number: (text: string) => number | null, names: (lcid: string) => Names | undefined): { text: Dialect; cell: Style } {
  const { decimal } = data
  const style: Style = {
    decimal,
    group: data.group,
    months: data.months,
    monthsShort: data.monthsShort,
    days: data.days,
    daysShort: data.daysShort,
    dateSeparator: '/',
    general: (n) => generalFormat(n).replace('.', decimal),
    bool: (b) => locale.booleans[b ? 'TRUE' : 'FALSE'],
    error: (code) => errorIn(code as ErrorCode, locale.errors),
    names,
    longDate: data.longDate,
    longTime: data.longTime,
    longMonths: data.longMonths,
    builtin: builtinCodes(data.builtin),
  }
  return { text: { codes: regionalCodes(data), style, number }, cell: { ...style, dateSeparator: data.dateSeparator } }
}
