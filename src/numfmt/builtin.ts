/**
 * Excel's built-in number formats by number (`numFmtId` 0 to 49, which an xlsx file stores without a code): the codes
 * Excel for Windows en-US shows them by (the display set `builtin`, `test/excel/win/display/`), as Excel for Mac but
 * the short date, `m/d/yyyy` (the Mac's `m/d/yy`). A locale shows some of them by its own codes (`Style.builtin`): the
 * short and medium dates and times, money, numbers with negatives in parentheses.
 */
import { localeOf } from '../locale/registry.js'
import type { Locale } from '../locale/types.js'

const SHORT_DATE = 'm/d/yyyy'
const LONG_TIME = 'h:mm:ss'

/** en-US's codes, by number; 23 to 26 are General, 27 to 36 (East Asian dates and times elsewhere) the short date or time. */
export const EN_BUILTIN: readonly string[] = [
  'General', '0', '0.00', '#,##0', '#,##0.00', '$#,##0_);($#,##0)', '$#,##0_);[Red]($#,##0)', '$#,##0.00_);($#,##0.00)', '$#,##0.00_);[Red]($#,##0.00)',
  '0%', '0.00%', '0.00E+00', '# ?/?', '# ??/??', SHORT_DATE, 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'h:mm AM/PM', 'h:mm:ss AM/PM', 'h:mm', LONG_TIME,
  `${SHORT_DATE} h:mm`, 'General', 'General', 'General', 'General', SHORT_DATE, SHORT_DATE, SHORT_DATE, SHORT_DATE, SHORT_DATE, LONG_TIME, LONG_TIME,
  LONG_TIME, LONG_TIME, SHORT_DATE, '#,##0_);(#,##0)', '#,##0_);[Red](#,##0)', '#,##0.00_);(#,##0.00)', '#,##0.00_);[Red](#,##0.00)',
  '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)', '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
  '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', 'mm:ss', '[h]:mm:ss', 'mm:ss.0', '##0.0E+0', '@',
]

/** The numbers of the built-in formats: 0 to 49. */
export const isBuiltin = (id: number) => Number.isInteger(id) && id >= 0 && id < EN_BUILTIN.length

/** The codes of a locale that shows built-in formats otherwise: its own by number, en-US's for the rest. */
export const builtinCodes =
  (own: Readonly<Record<number, string>>) =>
  (id: number): string =>
    own[id] ?? EN_BUILTIN[id]

/** The codes a locale (a loaded locale's tag or an object; en-US without one) shows the built-in formats by, by number: its own, en-US's for the rest. */
export const builtinIn = (where: string | Locale | undefined): ((id: number) => string) => builtinCodes(localeOf(where).regional?.builtin ?? {})
