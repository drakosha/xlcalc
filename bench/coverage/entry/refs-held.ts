/**
 * Held out from the set `refs`, predicted by the engine before Excel ran it: other cells and formats (column D, a
 * column E formatted as percent, a sheet U, a name), the difference of two formulas with formats (two dates, two
 * percents, a date and a time), the functions that pass a format on to new arguments (ROUNDDOWN and TRUNC by the
 * engine's guess, other rounding and statistical functions as giving none), the first cell of a range and of a whole
 * column, and cells of other formats typed into.
 */
import type { CellInput, NumberFormat } from '../../../src/contract.js'
import { type EntryCase, type EntryData, entryCases } from './cases.js'

const D: [number | string, NumberFormat | undefined][] = [
  [46096, 'd mmmm yyyy'],
  [0.3, 10],
  [100, 5],
  [0.5, '0.0%'],
  [0.25, 21],
  [2, undefined],
  [46096, 15],
  [3.5, 13],
  [1000, 4],
  [100000, 48],
  [46200, 14],
  [0.1, 9],
  [5, '[Red]0'],
  [NaN, undefined],
  [0.2, 9],
  [0.7, 46],
  [100, '€#,##0'],
]

/** D1 a long date of the book's own code, D2 0.00%, D3 money, D4 a percent of the book's own code, D5 h:mm:ss, D6 2 in General, D7 d-mmm-yy, D8 a fraction, D9 #,##0.00, D10 scientific, D11 the short date, D12 and D15 percents, D13 a code of the book's own, D14 blank, D16 [h]:mm, D17 euros; column E in a percent format, E5 0.4. */
const cells: Record<string, CellInput> = Object.fromEntries(D.flatMap(([v, f], i) => (Number.isNaN(v) ? [] : [[`D${i + 1}`, { value: v, ...(f === undefined ? {} : { numberFormat: f }) }]])))

export const data: EntryData = {
  cells: { ...cells, E5: { value: 0.4 } },
  columns: { E: { numberFormat: 9 } },
  sheets: [{ name: 'U', cells: { C1: { value: 46096, numberFormat: 14 }, C2: { value: 1 } } }],
  names: [{ name: 'Rate', ref: 'S!$D$12' }],
}

const GENERAL: Record<string, string[]> = {
  one: D.flatMap(([v, f], i) => (f === undefined || Number.isNaN(v) ? [] : [`=D${i + 1}`])),
  diff: ['=D11-D1', '=D11-D7', '=D12-D15', '=D3-D3', '=D11-D12', '=D12-D11', '=D5-D5', '=D11-D5', '=D5-D11', '=1-D11', '=D6-D11', '=D11-1', '=D11-TODAY()', '=NOW()-D11', '=TODAY()-D12', '=D12-TODAY()', '=D16-D5'],
  sum: ['=D6+D12', '=D12+D3', '=D3+D12', '=D14+D12', '=D6*D12', '=D12/100', '=-D3', '=--D12', '=(D6+D12)*1', '=IF(D6>0,D12)'],
  pass: ['=SUM(D12,D3)', '=SUM(D6,D3)', '=MAX(D6,D11)', '=MIN(D12)', '=AVERAGE(D6:D7)', '=AVERAGE(D7:D8)', '=MEDIAN(D3,D6)', '=ROUND(D6,D12)', '=ROUNDUP(D3,0)', '=INT(D11)', '=ROUNDDOWN(D12,1)', '=TRUNC(D12)'],
  other: ['=MROUND(D12,0.1)', '=CEILING(D3,1)', '=FLOOR(D3,1)', '=LARGE(D12:D15,1)', '=SMALL(D12:D15,1)', '=MODE(D12,D12)', '=SUMIF(D12:D15,">0")', '=AVERAGEIF(D12:D15,">0")', '=MAXA(D12)', '=SIGN(D12)', '=MOD(D12,1)', '=ABS(D3)', '=STDEV(D12,D15)', '=EDATE(D11,1)', '=AVERAGEA(D12)', '=DAY(D11)'],
  range: ['=SUM(D:D)', '=SUM(E:E)', '=E5', '=E7', '=SUM(D14:D15)', '=SUM(D6:D7)', '=SUM(D12:D15)', '=SUM(D15:D12)', '=SUM((D6,D12))'],
  sheet: ['=U!C1', "='U'!C1*2", '=SUM(U!C1:C2)', '=U!C2+D12'],
  name: ['=Rate', '=Rate+D6', '=D6+Rate'],
  ru: ['=СУММ(D12;D3)', '=ОКРУГЛВВЕРХ(D3;0)', '=ЦЕЛОЕ(D11)', '=МЕДИАНА(D3;D6)', '=ОТБР(D12)', '=СУММ(D:D)', '=ОКРУГЛВНИЗ(D12;1)'],
}

const FORMATTED: [string, NumberFormat, string[]][] = [
  ['in-sci', 11, ['=D12', '=D11']],
  ['in-frac', 12, ['=D12']],
  ['in-acct', 44, ['=D12', '=D6']],
  ['in-pct2', 10, ['=D12', '=D3']],
  ['in-grp', 3, ['=D12']],
  ['in-text', '@', ['=D12']],
  ['in-datel', 15, ['=D11', '=D5']],
  ['in-own', '[Red]0', ['=D11']],
]

export const buildEntryCases = (): EntryCase[] => entryCases(GENERAL, FORMATTED)
