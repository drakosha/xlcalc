/**
 * The ribbon's Number Format box over codes that tell whether Excel for Windows names a code with a bracketed symbol
 * Accounting by its first section alone or by any of its sections: the menu's accounting section after another section
 * and before anything, the paddings and the fill in other forms, no thousands separator, a colour or a condition before
 * it. Held out from the sets `ribbon-accounting`, `-held` and `-held-2`; predicted by `numberFormatCategory` before Excel ran it.
 */
import { collector, type UiCase } from './cases.js'

const CODES = [
  '#,##0.00;_-[$£-809]* #,##0.00_-',
  '[$£-809]#,##0.00;_-[$£-809]* #,##0.00_-',
  '"x";_-[$£-809]* #,##0.00_-;_-[$£-809]* "-"??_-;_-@_-',
  '_-[$£-809]* #,##0.00_-;"x"',
  '_-[$£-809]* #,##0.00_-;;;',
  '_-[$£-809]* #,##0.00_-;dd.mm.yyyy',
  '_([$£-809]* #,##0.00_-',
  '_)[$£-809]* #,##0.00_(',
  '[$£-809]* #,##0.00',
  '_-[$£-809]*x#,##0.00_-',
  '_-[$£-809]* 0.00_-',
  '[Red]_-[$£-809]* #,##0.00_-',
  '[>0]_-[$£-809]* #,##0.00_-;_-[$£-809]* "-"??_-',
  '_-* #,##0.00 [$€-407]_-',
  '_-* #,##0.00[$€-407]_-',
  '_-[$USD] * #,##0.00_-',
  '_ [$¥-804]* #,##0.000_ ',
]

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  const cells = Object.fromEntries(CODES.map((numberFormat, i) => [`B${i + 1}`, { value: -1234.5, numberFormat }]))
  add('held', { cells, act: { do: 'ribbon', cells: Object.keys(cells) } })
  return out
}
