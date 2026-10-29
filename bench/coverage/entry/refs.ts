/**
 * The format a typed formula takes from the cells it refers to, predicted by the engine before Excel ran it: a reference
 * to a cell of each kind of format, the operators and functions around a reference to a percent (C3) beside a number in
 * General (C1), two references of other formats, ranges whose first cell is General or blank, another sheet, a defined
 * name, and cells of a format other than General typed into (as a cell that took a format from an earlier formula).
 */
import type { CellInput, NumberFormat } from '../../../src/contract.js'
import { type EntryCase, type EntryData, entryCases } from './cases.js'

const C: [number, NumberFormat | undefined][] = [
  [2, undefined],
  [46096, 14],
  [0.25, 9],
  [0.125, 10],
  [1234.5, 7],
  [1234.5, '$#,##0.00'],
  [0.75, 20],
  [1234.5, 2],
  [1234.5, 3],
  [12345, 11],
  [1.5, 12],
  [5, '@'],
  [1234.5, 44],
  [3, '0.0 "kg"'],
  [46096.5, 22],
  [NaN, 9],
  [0.5, 9],
  [46100, 'yyyy-mm-dd'],
  [7, undefined],
]

/** C1 2 in General, C2 a date, C3 and C4 percents, C5 and C6 money, C7 a time, C8 and C9 numbers, C10 scientific, C11 a fraction, C12 text's format, C13 accounting, C14 a code of the book's own, C15 a date and time, C16 blank in a percent format, C17 a percent, C18 a date of the book's own code, C19 7 in General. */
const cells: Record<string, CellInput> = Object.fromEntries(C.map(([v, f], i) => [`C${i + 1}`, { ...(Number.isNaN(v) ? {} : { value: v }), ...(f === undefined ? {} : { numberFormat: f }) }]))

export const data: EntryData = {
  cells,
  sheets: [{ name: 'T', cells: { C1: { value: 0.5, numberFormat: 9 }, C2: { value: 3 } } }],
  names: [
    { name: 'Pct', ref: 'S!$C$3' },
    { name: 'Num', ref: 'S!$C$1' },
  ],
}

const GENERAL: Record<string, string[]> = {
  one: C.flatMap(([, f], i) => (f === undefined ? [] : [`=C${i + 1}`])),
  op: ['=C3+C1', '=C1+C3', '=C3-C1', '=C1-C3', '=C3-1', '=C3*C1', '=C1*C3', '=C3*2', '=2*C3', '=C3/C1', '=C1/C3', '=C3/2', '=C3^2', '=-C3', '=+C3', '=(C3)', '=C3%', '=C3&""', '=C3=C1', '=C3+C5', '=C5+C3', '=C2+C3', '=C3+C2', '=C2-C15', '=C5*2', '=C5/2', '=C5-C1', '=C1-C5', '=C7+C1', '=C1+C7', '=C2*1', '=C1+C2', '=TODAY()+C3', '=C3+TODAY()', '=C16+C1', '=C1+C16', '=C19+C3'],
  fn: ['=SUM(C3)', '=SUM(C3,C1)', '=SUM(C1,C3)', '=SUM(C3:C4)', '=SUM(C1:C3)', '=SUM(C3:C5)', '=SUM(C2:C3)', '=SUM(C16:C17)', '=SUM(C:C)', '=SUM(C3,C5)', '=SUM(C5,C3)', '=MAX(C3,C1)', '=MIN(C1,C3)', '=AVERAGE(C3,C17)', '=ROUND(C3,2)', '=ABS(C3)', '=INT(C5)', '=IF(C1>0,C3,0)', '=IF(TRUE,C3)', '=CONCAT(C3)', '=INDEX(C3:C4,1)', '=PRODUCT(C3,2)', '=SUMPRODUCT(C3:C4)', '=SUBTOTAL(9,C3:C4)', '=N(C3)', '=CHOOSE(1,C3)', '=MEDIAN(C3,C17)', '=COUNT(C3:C4)', '=ROUNDUP(C3,1)', '=SUM(C3)*2', '=SUM(C5)+C3'],
  fnru: ['=СУММ(C3)', '=СУММ(C1;C3)', '=СУММ(C3:C4)', '=ОКРУГЛ(C3;2)', '=ЕСЛИ(C1>0;C3;0)', '=МАКС(C3;C1)', '=СУММ(C1:C3)'],
  sheet: ['=T!C1', '=T!C1+C1', '=C1+T!C1', '=SUM(T!C1:C2)', '=T!C2'],
  name: ['=Pct', '=Pct*2', '=Pct+C1', '=SUM(Pct)', '=Num+Pct'],
}

const FORMATTED: [string, NumberFormat, string[]][] = [
  ['in-num', 2, ['=C3', '=C2', '=C5']],
  ['in-pct', 9, ['=C1', '=C2', '=C5', '=C4']],
  ['in-date', 14, ['=C3', '=C5', '=C1', '=C15']],
  ['in-cur', 7, ['=C3', '=C2']],
  ['in-own', '0.0 "kg"', ['=C3']],
  ['in-time', 20, ['=C2', '=C3']],
]

export const buildEntryCases = (): EntryCase[] => entryCases(GENERAL, FORMATTED)
