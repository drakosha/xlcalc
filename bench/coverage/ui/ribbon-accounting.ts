/**
 * The ribbon's Number Format box over the codes of the Accounting Number Format menu of Excel for Windows (`$` on the
 * Home tab), in en-US and ru-RU (`bench/windows/ribbon.ts`): each item's code in both regional formats, the built-in 44
 * the button's left part sets, by number and as its codes, and the menu's forms with another symbol, none, a symbol
 * written bare or quoted, no decimals, no text section. Predicted by `numberFormatCategory` before Excel ran it.
 */
import type { CellInput } from '../../../src/contract.js'
import { collector, type UiCase } from './cases.js'

const MENU = [
  '_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* "-"??_);_(@_)',
  '_-[$$-409]* #,##0.00_ ;_-[$$-409]* -#,##0.00 ;_-[$$-409]* "-"??_ ;_-@_ ',
  '_-* #,##0.00 [$₽-419]_-;-* #,##0.00 [$₽-419]_-;_-* "-"?? [$₽-419]_-;_-@_-',
  '_-[$£-809]* #,##0.00_-;-[$£-809]* #,##0.00_-;_-[$£-809]* "-"??_-;_-@_-',
  '_([$€-2] * #,##0.00_);_([$€-2] * (#,##0.00);_([$€-2] * "-"??_);_(@_)',
  '_-[$€-2] * #,##0.00_-;-[$€-2] * #,##0.00_-;_-[$€-2] * "-"??_-;_-@_-',
  '_ [$¥-804]* #,##0.00_ ;_ [$¥-804]* -#,##0.00_ ;_ [$¥-804]* "-"??_ ;_ @_ ',
  '_-* #,##0.00 [$CHF-100C]_-;-* #,##0.00 [$CHF-100C]_-;_-* "-"?? [$CHF-100C]_-;_-@_-',
]
const LEFT: CellInput['numberFormat'][] = [
  44,
  '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
  '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
  '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  '_-* #,##0.00 "₽"_-;-* #,##0.00 "₽"_-;_-* "-"?? "₽"_-;_-@_-',
  '_-* #,##0.00 ₽_-;-* #,##0.00 ₽_-;_-* "-"?? ₽_-;_-@_-',
]
const FORMS = [
  '_-[$$-409]* #,##0.00_-;-[$$-409]* #,##0.00_-;_-[$$-409]* "-"??_-;_-@_-',
  '_ [$€-2] * #,##0.00_ ;_ [$€-2] * -#,##0.00_ ;_ [$€-2] * "-"??_ ;_ @_ ',
  '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-',
  '_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ ',
  '_-£* #,##0.00_-;-£* #,##0.00_-;_-£* "-"??_-;_-@_-',
  '_-"£"* #,##0.00_-;-"£"* #,##0.00_-;_-"£"* "-"??_-;_-@_-',
  '_-[$£-809]* #,##0_-;-[$£-809]* #,##0_-;_-[$£-809]* "-"_-;_-@_-',
  '_ [$¥-804]* #,##0.00_ ;_ [$¥-804]* -#,##0.00_ ;_ [$¥-804]* "-"??_ ',
]

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  const cells = Object.fromEntries([...MENU, ...LEFT, ...FORMS].map((numberFormat, i) => [`B${i + 1}`, { value: -1234.5, numberFormat }]))
  add('menu', { cells, act: { do: 'ribbon', cells: Object.keys(cells) } })
  return out
}
