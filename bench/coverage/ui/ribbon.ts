/**
 * The ribbon's Number Format box (the Home tab) for each cell selected, against `numberFormatCategory`, and each item
 * of its list applied, against `numberFormatChoices`, predicted before Excel ran it. The built-in formats 0 to 49 by
 * number (as an xlsx file stores them, no code) and by their en-US codes written as codes of the file's own, ru-RU's
 * codes of the built-in formats, and variants of each category: decimals, group separators, negatives in red and in
 * parentheses, currency symbols of en-US and ru-RU in each place and form (`$`, `[$$-409]`, `"₽"`, `[$₽-419]`, `[$€-x]`),
 * accounting formats, dates and times in other orders and with locale tags (`[$-F800]`, `[$-F400]`), fractions,
 * exponents, text, General with text, hidden formats, colours and conditions, the special formats. A few formats on
 * text, blanks, negatives and dates (does the box look at the value). Each item applied to a cell in General and to
 * cells of another format already (does the item's code depend on it). Both locales run the same cases.
 */
import type { CellInput } from '../../../src/contract.js'
import { EN_BUILTIN } from '../../../src/numfmt/builtin.js'
import { ruRU } from '../../../src/locales/ru-RU.js'
import { collector, type UiCase } from './cases.js'

/** A case's cells at most: each takes a selection and a read of the box, under the step's time for a case. */
const PER_CASE = 25

const NUMBER = [
  '0.000', '0.0', '#,##0.0', '#,##0.000', '0.0000000000', '0.000000000000000000000000000000', '0.00_);(0.00)', '0.00_);[Red](0.00)', '0.00;[Red]0.00',
  '0.00;[Red]-0.00', '#,##0.00;[Red]#,##0.00', '#,##0.00;[Red]-#,##0.00', '#,##0;-#,##0', '0;[Red]0', '[Red]0.00', '[Blue]0.00', '0.00;[Blue]-0.00',
  '0.00;-0.00', '0.00;-0.00;0.00', '0.00;-0.00;"-"', '#.##', '#,###', '00.00', '0.00 "kg"', '#,##0.00 ', '# ##0.00', '0.00;(0.00)', '(0.00)',
]
const PERCENT = ['0.0%', '0.000%', '#,##0.00%', '0.00%;[Red]-0.00%', '0%;-0%', '0.0000000000%']
const FRACTION = ['# ???/???', '?/?', '??/??', '# ?/2', '# ?/4', '# ?/8', '# ??/16', '# ?/10', '# ??/100', '0 ?/?', '# ?/3', '#" "?/?']
const SCIENTIFIC = ['0.000E+00', '0E+00', '0.0E+00', '0.0E+0', '0.00E-00', '#.##E+00', '0.00E+000', '0.0000000000E+00']
// Excel repairs a file with `@;@`, a text section first, so the set leaves it out
const TEXT = ['"x"@', '@" kg"', ';;;@', '0;-0;0;@', 'General;[Red]-General', '[Red]General', 'General" kg"', '0;-0;0;General', 'General;General;General;@', ';;;', ';;', '0;;']
const CURRENCY = [
  '$#,##0.00', '$#,##0', '$#,##0.000', '$#,##0.00;[Red]$#,##0.00', '$#,##0.00;-$#,##0.00', '$#,##0.00;[Red]-$#,##0.00', '"$"#,##0.00', '\\$#,##0.00',
  '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)', '[$$-409]#,##0.00', '[$$-en-US]#,##0.00', '[$$-409]#,##0.00;[Red]-[$$-409]#,##0.00', '$0.00', '#,##0.00 $', '$ #,##0.00',
  '[$€-x-euro2] #,##0.00', '[$€-407] #,##0.00', '#,##0.00 [$€-407]', '[$€-2] #,##0.00', '#,##0.00 [$€-1]', '€#,##0.00', '#,##0.00 €', '[$£-809]#,##0.00', '£#,##0.00',
  '[$¥-411]#,##0', '¥#,##0', '[$USD] #,##0.00', '[$CHF] #,##0.00', '[$EUR] #,##0.00',
  '#,##0.00 [$₽-419]', '#,##0.00 "₽"', '#,##0.00 ₽', '#,##0.00\\ "₽"', '#,##0.00\\ [$₽-419]', '#,##0.00 [$₽-ru-RU]', '#,##0 "₽"', '#,##0.00 "₽";[Red]-#,##0.00 "₽"',
  '#,##0.00"р."', '#,##0.00 "руб."', '[$₽-419]#,##0.00', '#,##0.00 [$RUB]',
]
const ACCOUNTING = [
  '_($* #,##0.000_);_($* (#,##0.000);_($* "-"???_);_(@_)', '_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* "-"??_);_(@_)',
  '_-[$$-409]* #,##0.00_ ;_-[$$-409]* -#,##0.00 ;_-[$$-409]* "-"??_ ;_-@_ ', '_([$€-x-euro2] * #,##0.00_);_([$€-x-euro2] * (#,##0.00);_([$€-x-euro2] * "-"??_);_(@_)',
  '_-* #,##0.00 [$€-407]_-;-* #,##0.00 [$€-407]_-;_-* "-"?? [$€-407]_-;_-@_-', '_-* #,##0.00 [$₽-419]_-;-* #,##0.00 [$₽-419]_-;_-* "-"?? [$₽-419]_-;_-@_-',
  '_-* #,##0.00\\ [$₽-419]_-;\\-* #,##0.00\\ [$₽-419]_-;_-* "-"??\\ [$₽-419]_-;_-@_-', '_-* #,##0.00\\ "₽"_-;\\-* #,##0.00\\ "₽"_-;_-* "-"??\\ "₽"_-;_-@_-',
  '_($* #,##0.00_)', '$* #,##0.00',
]
const DATE = [
  'm/d/yy', 'mm/dd/yyyy', 'mm/dd/yy', 'm/d', 'd/m/yyyy', 'dd/mm/yyyy', 'yyyy-mm-dd', 'd.m.yy', 'dd.mm.yy', 'd.m.yyyy', 'mmmm-yy', 'mmmm d, yyyy', 'mmmmm',
  'mmmmm-yy', 'd-mmm-yyyy', 'dddd, mmmm d, yyyy', 'dddd, mmmm dd, yyyy', '[$-F800]dddd, mmmm dd, yyyy', '[$-F800]dddd, mmmm d, yyyy', '[$-F800]dddd\\, mmmm dd\\, yyyy',
  'd mmmm yyyy', 'd mmmm yyyy "г."', 'dddd, d mmmm yyyy "г."', 'mmmm yyyy', 'yyyy', 'mmm d', 'dd/mmm/yy', '[$-409]m/d/yy', '[$-409]mmmm d, yyyy;@', '[$-419]dd.mm.yyyy',
  '[$-x-sysdate]dddd, mmmm dd, yyyy', 'm/d/yy h:mm', 'm/d/yy h:mm AM/PM', 'dd.mm.yy h:mm', 'yyyy-mm-dd hh:mm:ss', 'dd/mm/yyyy;@', 'm/d/yyyy;@', 'dd.mm.yyyy;@',
  'dddd', 'ddd, mmm d', 'm/d/yy;[Red]m/d/yy', '[Red]m/d/yy', 'yyyy\\-mm\\-dd', 'd/m/yy', 'yy/m/d',
]
const TIME = [
  'hh:mm', 'hh:mm:ss', 'hh:mm AM/PM', '[$-F400]h:mm:ss AM/PM', '[$-F400]h:mm:ss', '[h]:mm', '[mm]:ss', 'h:mm:ss.000', 'h"h" mm"m"', '[$-409]h:mm:ss AM/PM;@',
  'h:mm;@', 'h:mm A/P', 'ss', 'h', '[$-419]h:mm', 'h:mm:ss;@', '[$-F400]h:mm:ss AM/PM;@',
]
const SPECIAL = ['00000', '00000-0000', '[<=9999999]###-####;(###) ###-####', '000-00-0000', '000000']
const OTHER = ['[Color10]0.00', '[>=100]0.00;[Red]0.00', '[<0]"neg";"pos"', '0.00;[Red]-0.00;"zero"', '[Red][<0]0.00;0.00', '[>100][Red]0;0', '"text"', '0" pcs"', '"$"0.00']

/** ru-RU's codes of the built-in formats that are not en-US's, once each. */
const RU_CODES = [...new Set(Object.values(ruRU.regional!.builtin))].filter((code) => !EN_BUILTIN.includes(code))

/** The variants, each once and none a built-in format's en-US code (the cases by number and code have those). */
const VARIANTS = [...new Set([...NUMBER, ...PERCENT, ...FRACTION, ...SCIENTIFIC, ...TEXT, ...CURRENCY, ...ACCOUNTING, ...DATE, ...TIME, ...SPECIAL, ...OTHER])].filter((code) => !EN_BUILTIN.includes(code))

/** Formats on other values than 1.5: text, a blank, a negative number, a date. */
const ON_VALUES = ['0.00', '$#,##0.00', 'm/d/yy', 'h:mm', '0%', '@', 'General', '0.000']
const VALUES: (CellInput['value'] | undefined)[] = ['x', undefined, -2.5, 45688]

/** Cells of other formats, each with the item applied to it: does the item's code depend on what the cell had. */
const BEFORE: [CellInput, number][] = [
  [{ value: 1.5, numberFormat: '0.000' }, 2],
  [{ value: 1.5, numberFormat: '$#,##0.000' }, 3],
  [{ value: 1.5, numberFormat: '[$€-x-euro2] #,##0.00' }, 3],
  [{ value: 1.5, numberFormat: '#,##0.00 [$₽-419]' }, 3],
  [{ value: 1.5, numberFormat: '_([$€-x-euro2] * #,##0.00_);_([$€-x-euro2] * (#,##0.00);_([$€-x-euro2] * "-"??_);_(@_)' }, 4],
  [{ value: 1.5, numberFormat: '0.000%' }, 8],
  [{ value: 1.5, numberFormat: 'dd/mm/yyyy' }, 5],
  [{ value: 1.5, numberFormat: 'd mmmm yyyy' }, 6],
  [{ value: 1.5, numberFormat: 'hh:mm' }, 7],
  [{ value: 1.5, numberFormat: '0.0E+00' }, 10],
  [{ value: 1.5, numberFormat: '# ???/???' }, 9],
  [{ value: 'x' }, 2],
  [{}, 3],
  [{ value: 1.5, numberFormat: '#,##0.0' }, 2],
  [{ value: 'x', numberFormat: '@' }, 2],
]

const cell = (i: number) => `B${i + 1}`

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  /** Cases of the ribbon's box over the cells given, PER_CASE at most each, the `own` ones' codes the file's own. */
  const boxes = (group: string, list: CellInput[], own = false) => {
    for (let at = 0; at < list.length; at += PER_CASE) {
      const part = list.slice(at, at + PER_CASE)
      const cells = Object.fromEntries(part.map((x, i) => [cell(i), x]))
      add(group, { cells, act: { do: 'ribbon', cells: Object.keys(cells) }, ...(own ? { own: Object.keys(cells) } : {}) })
    }
  }
  boxes('builtin', EN_BUILTIN.map((_, id) => ({ value: 1.5, numberFormat: id })))
  boxes('code', EN_BUILTIN.map((code) => ({ value: 1.5, numberFormat: code })), true)
  boxes('ru-code', RU_CODES.map((code) => ({ value: 1.5, numberFormat: code })), true)
  boxes('variant', VARIANTS.map((code) => ({ value: 1.5, numberFormat: code })))
  boxes('value', ON_VALUES.flatMap((code) => VALUES.map((value) => ({ ...(value === undefined ? {} : { value }), numberFormat: code }))))
  const items = Array.from({ length: 11 }, (_, i) => i + 1)
  add('choose', { cells: Object.fromEntries(items.map((_, i) => [cell(i), { value: 1.5 }])), act: { do: 'choose', picks: items.map((k, i) => [cell(i), k]) } })
  add('choose', { cells: Object.fromEntries(BEFORE.map(([x], i) => [cell(i), x])), act: { do: 'choose', picks: BEFORE.map(([, k], i) => [cell(i), k]) } })
  return out
}
