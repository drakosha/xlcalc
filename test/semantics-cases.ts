/**
 * Semantics cases for the whitelist (SUM, SUMIF, SUMIFS, COUNTIFS, IF, IFERROR, XLOOKUP, ROUND, EOMONTH,
 * DATE, TEXT): each case with the answer per Excel documentation; the LibreOffice oracle confirms it.
 */
import type { BookInput, CellInput, CellValue } from '../src/contract.js'
import { dateToSerial } from '../src/model.js'
import { err } from './formula-values.js'

export const SEM = 'Sem'
const A = ['Аренда', 'аренда', 'Арендй', 'Арeнда*', 'x*y', 'Чай', 'Чаи'] // 'Арeнда*' has a Latin e
const B = [1, 2, 4, 8, 16, 32, 64]
const C = ['2024-12-31', '2025-01-01', '2025-01-02', '2025-02-01', '2025-03-01', '2024-06-30', '2025-12-31']
const D: CellInput['value'][] = [1, '2', true]

export function semanticsBook(cases: SemCase[]): BookInput {
  const sem: Record<string, CellInput> = {}
  A.forEach((v, i) => (sem[`A${i + 1}`] = { value: v }))
  B.forEach((v, i) => (sem[`B${i + 1}`] = { value: v }))
  C.forEach((v, i) => (sem[`C${i + 1}`] = { value: { date: v } }))
  D.forEach((v, i) => (sem[`D${i + 1}`] = { value: v }))
  const cs: Record<string, CellInput> = {}
  cases.forEach((c, i) => (cs[`A${i + 1}`] = { formula: c.formula }))
  return { sheets: [{ name: SEM, cells: sem }, { name: 'Cases', cells: cs }] }
}

const s = dateToSerial

export interface SemCase {
  id: string
  fn: string
  formula: string
  /** The answer per Excel documentation. */
  excel: CellValue
  note?: string
}


export const CASES: SemCase[] = [
  { id: 's1', fn: 'SUMIFS', formula: 'SUMIFS(Sem!B1:B7,Sem!C1:C7,">=2025-01-01")', excel: 94, note: 'date as text in the criterion' },
  { id: 's2', fn: 'SUMIFS', formula: 'SUMIFS(Sem!B1:B7,Sem!C1:C7,">="&DATE(2025,1,1))', excel: 94, note: 'date through DATE' },
  { id: 's3', fn: 'SUMIFS', formula: 'SUMIFS(Sem!B1:B7,Sem!A1:A7,"Ар*")', excel: 15, note: 'wildcard *' },
  { id: 's4', fn: 'SUMIFS', formula: 'SUMIFS(Sem!B1:B7,Sem!A1:A7,"x~*y")', excel: 16, note: 'escaped ~*' },
  { id: 's5', fn: 'SUMIFS', formula: 'SUMIFS(Sem!B1:B7,Sem!A1:A7,"аренда")', excel: 3, note: 'Cyrillic, case-insensitive' },
  { id: 's6', fn: 'SUMIFS', formula: 'SUMIFS(Sem!B1:B7,Sem!A1:A7,"Чай")', excel: 32, note: '«й» ≠ «и»' },
  { id: 's7', fn: 'SUMIFS', formula: 'SUMIFS(Sem!B1:B7,Sem!A1:A7,"<>Аренда")', excel: 124, note: 'not equal, case-insensitive' },
  { id: 's8', fn: 'SUMIFS', formula: 'SUMIFS(Sem!B1:B7,Sem!A1:A7,"Ча?")', excel: 96, note: 'wildcard ?' },
  { id: 'c1', fn: 'COUNTIFS', formula: 'COUNTIFS(Sem!C1:C7,">=2025-01-01")', excel: 5, note: 'date as text' },
  { id: 'c2', fn: 'COUNTIFS', formula: 'COUNTIFS(Sem!A1:A7,"Ар*",Sem!B1:B7,">2")', excel: 2, note: 'wildcard and number' },
  { id: 'c3', fn: 'COUNTIFS', formula: 'COUNTIFS(Sem!A1:A7,"чаи")', excel: 1, note: 'case, «й» ≠ «и»' },
  { id: 'x1', fn: 'XLOOKUP', formula: 'XLOOKUP("Чаи",Sem!A1:A7,Sem!B1:B7)', excel: 64, note: 'exact match' },
  { id: 'x2', fn: 'XLOOKUP', formula: 'XLOOKUP("нет",Sem!A1:A7,Sem!B1:B7)', excel: err('NA'), note: 'not found' },
  { id: 'x3', fn: 'XLOOKUP', formula: 'XLOOKUP("нет",Sem!A1:A7,Sem!B1:B7,"пусто")', excel: 'пусто', note: 'if_not_found' },
  { id: 'x4', fn: 'XLOOKUP', formula: 'XLOOKUP("Ар*",Sem!A1:A7,Sem!B1:B7,0,2)', excel: 1, note: 'wildcard, match_mode 2' },
  { id: 'x5', fn: 'XLOOKUP', formula: 'XLOOKUP("аренда",Sem!A1:A7,Sem!B1:B7,0,0,-1)', excel: 2, note: 'from the end, case-insensitive' },
  { id: 'r1', fn: 'ROUND', formula: 'ROUND(2.5,0)', excel: 3 },
  { id: 'r2', fn: 'ROUND', formula: 'ROUND(-2.5,0)', excel: -3 },
  { id: 'r3', fn: 'ROUND', formula: 'ROUND(1.005,2)', excel: 1.01 },
  { id: 'r4', fn: 'ROUND', formula: 'ROUND(2.675,2)', excel: 2.68 },
  { id: 'r5', fn: 'ROUND', formula: 'ROUND(1234.5678,-2)', excel: 1200 },
  { id: 'e1', fn: 'EOMONTH', formula: 'EOMONTH(DATE(2025,1,31),1)', excel: s('2025-02-28') },
  { id: 'e2', fn: 'EOMONTH', formula: 'EOMONTH(DATE(2024,1,15),1)', excel: s('2024-02-29'), note: 'leap year' },
  { id: 'e3', fn: 'EOMONTH', formula: 'EOMONTH(DATE(2025,3,15),-1)', excel: s('2025-02-28') },
  { id: 'e4', fn: 'EOMONTH', formula: 'EOMONTH(Sem!C2,0)', excel: s('2025-01-31') },
  { id: 'd1', fn: 'DATE', formula: 'DATE(2025,1,1)', excel: 45658 },
  { id: 'd2', fn: 'DATE', formula: 'DATE(2025,13,1)', excel: 46023, note: 'month out of range' },
  { id: 'd3', fn: 'DATE', formula: 'DATE(2025,1,0)', excel: 45657, note: 'day 0' },
  { id: 'd4', fn: 'DATE', formula: 'DATE(1900,3,1)', excel: 61, note: 'phantom 1900-02-29' },
  { id: 'd5', fn: 'DATE', formula: 'DATE(1900,2,28)', excel: 59 },
  { id: 't1', fn: 'TEXT', formula: 'TEXT(1234.5,"0.00")', excel: '1234.50' },
  { id: 't2', fn: 'TEXT', formula: 'TEXT(1234567.891,"#,##0.00")', excel: '1,234,567.89' },
  { id: 't3', fn: 'TEXT', formula: 'TEXT(0.256,"0%")', excel: '26%' },
  { id: 't4', fn: 'TEXT', formula: 'TEXT(DATE(2025,1,1),"YYYY-MM-DD")', excel: '2025-01-01', note: 'date' },
  { id: 't5', fn: 'TEXT', formula: 'TEXT(DATE(2025,3,5),"DD.MM.YYYY")', excel: '05.03.2025', note: 'date' },
  { id: 't6', fn: 'TEXT', formula: 'TEXT(-1234.5,"0")', excel: '-1235' },
  { id: 'u1', fn: 'SUM', formula: 'SUM(Sem!D1:D3)', excel: 1, note: 'text and booleans in a range do not count' },
  { id: 'u2', fn: 'SUM', formula: 'SUM("2",1)', excel: 3, note: 'numeric text as an argument' },
  { id: 'u3', fn: 'SUM', formula: 'SUM(TRUE,1)', excel: 2, note: 'boolean as an argument' },
  { id: 'u4', fn: 'SUM', formula: 'SUM(Sem!B1:B7)', excel: 127 },
  { id: 'u5', fn: 'SUM', formula: 'SUM(1,"abc")', excel: err('VALUE'), note: 'non-numeric text as an argument' },
  { id: 'i1', fn: 'IF', formula: 'IF(1>2,"a","b")', excel: 'b' },
  { id: 'i2', fn: 'IF', formula: 'IF("abc",1,2)', excel: err('VALUE'), note: 'text as the condition' },
  { id: 'i3', fn: 'IF', formula: 'IF(0,1)', excel: false },
  { id: 'i4', fn: 'IF', formula: 'IF(TRUE,1/0,2)', excel: err('DIV0') },
  { id: 'i5', fn: 'IF', formula: 'IF(FALSE,1/0,2)', excel: 2, note: 'error in the branch not taken' },
  { id: 'i6', fn: 'IF', formula: 'IF(TRUE,"он сказал ""да""","")', excel: 'он сказал "да"', note: 'quotes in a literal' },
  { id: 'f1', fn: 'IFERROR', formula: 'IFERROR(1/0,"x")', excel: 'x' },
  { id: 'f2', fn: 'IFERROR', formula: 'IFERROR(1,"x")', excel: 1 },
  { id: 'f3', fn: 'IFERROR', formula: 'IFERROR(Sem!A99,"x")', excel: 0, note: 'empty cell' },
  { id: 'm1', fn: 'SUMIF', formula: 'SUMIF(Sem!A1:A7,"Ар*",Sem!B1:B7)', excel: 15, note: 'wildcard' },
  { id: 'm2', fn: 'SUMIF', formula: 'SUMIF(Sem!B1:B7,">10")', excel: 112, note: 'without a sum range' },
  { id: 'm3', fn: 'SUMIF', formula: 'SUMIF(Sem!C1:C7,">=2025-01-01",Sem!B1)', excel: 94, note: 'sum range from its top-left cell' },
  { id: 'x6', fn: 'XLOOKUP', formula: 'XLOOKUP(10,Sem!B1:B7,Sem!A1:A7,"нет",-1)', excel: 'Арeнда*', note: 'exact or next smaller' },
  { id: 'x7', fn: 'XLOOKUP', formula: 'XLOOKUP(10,Sem!B1:B7,Sem!A1:A7,"нет",1)', excel: 'x*y', note: 'exact or next larger' },
  { id: 't7', fn: 'TEXT', formula: 'TEXT(-1234.5,"0.00;(0.00)")', excel: '(1234.50)', note: 'negative section' },
  { id: 't8', fn: 'TEXT', formula: 'TEXT(0.5,"#.##")', excel: '.5' },
  { id: 't9', fn: 'TEXT', formula: 'TEXT(1234567,"#,##0,")', excel: '1,235', note: 'scaling by 1000' },
  { id: 't10', fn: 'TEXT', formula: 'TEXT(45658.75,"dd mmm yyyy hh:mm:ss")', excel: '01 Jan 2025 18:00:00' },
  { id: 't11', fn: 'TEXT', formula: 'TEXT(45658.75,"h:mm AM/PM")', excel: '6:00 PM' },
  { id: 't12', fn: 'TEXT', formula: 'TEXT(45658,"mmmm d, yyyy")', excel: 'January 1, 2025' },
  { id: 't13', fn: 'TEXT', formula: 'TEXT(1234.5,"""$""#,##0.00")', excel: '$1,234.50', note: 'quoted literal' },
  { id: 't14', fn: 'TEXT', formula: 'TEXT(12,"000-00")', excel: '000-12', note: 'literal between digits' },
  { id: 't15', fn: 'TEXT', formula: 'TEXT("abc","0.00")', excel: 'abc', note: 'text as is' },
  { id: 't16', fn: 'TEXT', formula: 'TEXT(0.05,"0.0%")', excel: '5.0%' },
]
