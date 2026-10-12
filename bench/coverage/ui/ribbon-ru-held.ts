/**
 * The ribbon's Number Format box over codes held out from the set `ribbon`, for Excel for Windows in ru-RU
 * (`bench/windows/ribbon.ts`): what its Format Cells window may list in ru-RU, written in en-US syntax as a file stores
 * them. Dates and times with `;@` and the locale's tags (`[$-419]`, `[$-FC19]`), numeric dates day first, the ruble
 * bracketed, quoted and bare in more forms, accounting without a symbol, Number's negatives in ru-RU's forms, fractions
 * with a quoted space, more decimals of percent and exponents. Predicted by `numberFormatCategory` before Excel ran it.
 */
import type { CellInput } from '../../../src/contract.js'
import { collector, type UiCase } from './cases.js'

const PER_CASE = 25

const DATE = [
  'dd/mm/yy;@', 'dd/mm;@', 'd/m/yy;@', 'd/m;@', 'yyyy-mm-dd;@', 'dd/mm/yy h:mm;@', 'dd/mm/yyyy h:mm;@', 'd/m/yyyy', 'dd/mm/yy',
  '[$-419]d mmm;@', '[$-419]d mmm yy;@', '[$-419]dd mmm yy;@', '[$-419]mmm yy;@', '[$-419]mmmm yyyy;@', '[$-419]d mmmm yyyy;@', '[$-419]mmmmm;@',
  '[$-419]mmmmm yy;@', '[$-419]d-mmm-yy;@', '[$-FC19]d mmmm yyyy "г.";@', '[$-FC19]dd mmmm yyyy "г.";@', '[$-FC19]dd mmmm;@', '[$-419]dddd, d mmmm yyyy;@',
  'd mmm;@', 'mmmm yyyy;@',
]
const TIME = ['[$-409]h:mm AM/PM;@', '[$-409]h:mm:ss AM/PM;@', 'mm:ss.0;@', '[h]:mm:ss;@', 'h:mm AM/PM;@', 'hh:mm;@', '[$-419]h:mm;@']
const MONEY = [
  '#,##0.00 [$₽-419];-#,##0.00 [$₽-419]', '#,##0 [$₽-419]', '#,##0.00\\ "₽";[Red]\\-#,##0.00\\ "₽"', '#,##0.00\\ ₽', '#,##0\\ ₽', '[$₽-419]\\ #,##0.00',
  '#,##0.00\\ [$$-409]', '#,##0.00\\ [$€-1]', '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-', '_-* #,##0.000_-;\\-* #,##0.000_-;_-* "-"???_-;_-@_-',
  '_-* #,##0\\ [$₽-419]_-;\\-* #,##0\\ [$₽-419]_-;_-* "-"\\ [$₽-419]_-;_-@_-',
]
const OTHER = ['0.00;[Red]\\-0.00', '#,##0;[Red]#,##0', '#,##0.000;[Red]#,##0.000', '#" "?/2', '#" "?/4', '#" "???/???', '0.0%', '0.000E+00']

const cell = (i: number) => `B${i + 1}`

export function buildUiCases(): UiCase[] {
  const { out, add } = collector()
  const list: CellInput[] = [...DATE, ...TIME, ...MONEY, ...OTHER].map((code) => ({ value: 1.5, numberFormat: code }))
  for (let at = 0; at < list.length; at += PER_CASE) {
    const cells = Object.fromEntries(list.slice(at, at + PER_CASE).map((x, i) => [cell(i), x]))
    add('held', { cells, act: { do: 'ribbon', cells: Object.keys(cells) } })
  }
  return out
}
