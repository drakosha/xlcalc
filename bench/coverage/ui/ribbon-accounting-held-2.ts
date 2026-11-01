/**
 * The ribbon's Number Format box over codes held out from the sets `ribbon-accounting` and `ribbon-accounting-held`,
 * after the rule took a bracketed symbol's forms without their zero and text sections: such forms with two sections
 * and one, en-US's quoted dollar and ru-RU's quoted ruble without the text's section, a zero section left empty, another
 * text section, negatives in red, more question marks than decimals. Predicted by `numberFormatCategory` before Excel ran it.
 */
import { collector, type UiCase } from './cases.js'

const CODES = [
  '_-[$£-809]* #,##0.00_-',
  '_([$$-409]* #,##0.00_)',
  '_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00)',
  '_-[$$-409]* #,##0.00_ ;_-[$$-409]* -#,##0.00 ',
  '_-* #,##0.00 [$₽-419]_-;-* #,##0.00 [$₽-419]_-',
  '_ [$¥-804]* #,##0.00_ ;_ [$¥-804]* -#,##0.00_ ',
  '_-[$€-2] * #,##0.00_-;-[$€-2] * #,##0.00_-',
  '_-* #,##0 [$CHF-100C]_-;-* #,##0 [$CHF-100C]_-',
  '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_)',
  '_-* #,##0.00 "₽"_-;-* #,##0.00 "₽"_-',
  '_-[$£-809]* #,##0.00_-;-[$£-809]* #,##0.00_-;_-[$£-809]* "-"??_-;@',
  '_-[$£-809]* #,##0.00_-;-[$£-809]* #,##0.00_-;;_-@_-',
  '_-[$£-809]* #,##0.00_-;[Red]-[$£-809]* #,##0.00_-;_-[$£-809]* "-"??_-;_-@_-',
  '_-[$£-809]* #,##0.00_-;-[$£-809]* #,##0.00_-;_-[$£-809]* "-"???_-;_-@_-',
]

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  const cells = Object.fromEntries(CODES.map((numberFormat, i) => [`B${i + 1}`, { value: -1234.5, numberFormat }]))
  add('held', { cells, act: { do: 'ribbon', cells: Object.keys(cells) } })
  return out
}
