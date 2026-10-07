/**
 * Number format codes by kind, as an xlsx file stores them (en-US syntax whatever the locale), each on the values
 * of its kind in a wide column, so that the column's width plays no part: numbers and their sections, scaling and
 * percents, fractions, exponents, currencies and locale tags, conditions, colours, padding and fills, General,
 * text sections, dates and times, elapsed time. The same book is shown by Excel en-US and Excel ru-RU. Excel does
 * not load a book with a lower-case exponent (`0.00e+00`), so there is none here.
 */
import { collector } from './cases.js'

const NUMBERS = [1234.567, -1234.567, 0, 0.5, -0.5, 1234567.891, 0.000123, 1e15, -1e-7, 'abc', true, '=NA()']

const NUMBER_CODES = [
  '0', '0.00', '#,##0', '#,##0.00', '#,##0.000', '0.0,', '#,##0,', '#,##0,,"M"', '0.0,,', '#', '#.##', '#.#', '0.#', '?.??', '??.??', '00000',
  '000-00-0000', '#,###', '0%', '0.0%', '0.00%', '%0', '0%%', '$#,##0.00', '$#,##0.00_);($#,##0.00)', '$#,##0.00_);[Red]($#,##0.00)', '#,##0;-#,##0',
  '0;-0;"zero"', '0;(0);"-"', '0;;', ';;;', '0;-0;;@', '+0;-0;0', '0.00;-0.00;0.00;"text: "@', '0 "pcs"', '"USD "0.00', '\\$0', '0\\%', '(0)', '0-0',
  '#,##0.00 [$₽-419]', '[$₽-419]#,##0.00', '#,##0.00 ₽', '[$$-409]#,##0.00', '[$€-2] #,##0.00', '[$-409]0.00', '0.00E+00', '0.0E+0', '##0.0E+0',
  '0.00E-00', '0E+00', '#E+0', '# ?/?', '# ??/??', '?/?', '# ?/8', '# ??/16', '0 ?/100', '#,##0 ?/?',
  '[>100]0.0;[<=-100]-0;0', '[>=1000000]0.0,,"M";[>=1000]0.0,"K";0', '[<0]"neg";[>0]"pos";"zero"', '[=1]"one";"other"', '[>5]0', '[Red][<0]0;0',
  '[<=9999999]###-####;(###) ###-####', '[>100]"big"', '[Red]0', '[Blue]0;[Red]-0', '[Green]0', '[Color10]0', '[COLOR10]0', '[Yellow]0;[Cyan]-0;[Magenta]0',
  '[White]0', '[Black]0', '[Color1]0', '[Color56]0', '[red]0', '[Red]General', '[Red]@', '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
  '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', '0_)', '0_W', 'General', 'General" pcs"', '"x"General', 'general', '@', '"Name: "@', '@" ok"', '@@',
]

const DATES = [45688.75, 0, 0.5, 60, 61, 1.25, 2958465.99999, 2958466, -1, 45688.999999, 1 / 86400, 0.75 / 86400, 45688.5003, 100.25, 'abc']

const DATE_CODES = [
  'm/d/yyyy', 'mm/dd/yy', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'mmmm d, yyyy', 'dddd, mmmm d, yyyy', 'ddd', 'mmmmm', 'yyyy-mm-dd hh:mm:ss', 'h:mm AM/PM',
  'h:mm:ss AM/PM', 'h:mm a/p', 'hh:mm', 'h:mm:ss.0', 'mm:ss.00', 'mm:ss.000', '[h]:mm:ss', '[h]:mm', '[mm]:ss', '[ss]', '[hh]:mm', '[h]', 'yy', 'yyyy',
  'm', 'd', 'h', 's', 'm/d/yyyy h:mm', 'dd/mm/yyyy', 'yyyy"年"m"月"d"日"', 'dd.mm.yyyy', '[$-409]mmmm', '[$-419]mmmm', 'mmmm', 'dddd', '[Red]m/d/yyyy',
]

export function buildDisplayCases() {
  const { out, add } = collector()
  for (const f of NUMBER_CODES) for (const v of NUMBERS) add('number', v, f, { width: 60 })
  for (const f of DATE_CODES) for (const v of DATES) add('date', v, f, { width: 60 })
  return out
}
