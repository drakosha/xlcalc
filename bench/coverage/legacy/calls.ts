/**
 * When Excel 365 puts `@` before a call in a formula of the old kind (`run.ts legacy calls`): calls whose answer is
 * one value or several depending on their arguments (INDEX, IF, CHOOSE, IFERROR, LET, the lookups, ROW and COLUMN),
 * array constants and functions with an array answer where a reference is taken (inside SUM), and the same at the
 * top of the formula. With `params` and `nesting` these are the rules of `fromLegacy`.
 */
import type { Probe } from '../probes.js'
import { probe } from './params.js'

const R = 'D!B1:B3'
const M = 'D!B1:C3'

const GROUPS: Record<string, string[]> = {
  index: [`INDEX(${M},2,2)`, `INDEX(${M},0,1)`, `INDEX(${M},2,0)`, `INDEX(${R},2)`, `INDEX(${M},2)`, `INDEX(D!B1:D1,2)`, `INDEX(${M},D!A12,1)`, `INDEX(${M},2,D!A12)`, `INDEX(${M},"2",1)`, `INDEX(${M},-1,1)`, `INDEX(${M},2.5,1)`, `INDEX(${M},TRUE,1)`, `INDEX(${M},1+1,1)`, `INDEX(${M},2,2,1)`, `INDEX(D!B:B,2)`, `INDEX({1,2;3,4},2,1)`, `INDEX({1,2;3,4},0,1)`, `INDEX({1,2},2)`, `INDEX(D!B1,1,1)`, `INDEX(${M},,1)`, `INDEX(${M},2,)`, `ABS(INDEX(${M},2,2))`, `-INDEX(${M},0,1)`, `INDEX(${M},2,2)+INDEX(${M},0,1)`],
  if: ['IF(TRUE,1,2)', 'IF(TRUE,D!B1)', `IF(TRUE,${R})`, 'IF(TRUE,{1,2})', `IF(TRUE,1,${R})`, `IF(D!A4,${R},D!C1:C3)`, 'IF(TRUE,SEQUENCE(2))', `IF(TRUE,INDEX(${M},2,2))`, `IF(TRUE,INDEX(${M},0,1))`, `IF(TRUE,IF(TRUE,${R}))`, `ABS(IF(TRUE,${R}))`],
  choose: ['CHOOSE(2,1,2)', 'CHOOSE(1,D!B1)', `CHOOSE(1,${R})`, 'CHOOSE(1,{1,2})', 'CHOOSE(D!A4,1,2)', `CHOOSE(2,1,${R})`],
  iferror: [`IFERROR(${R},0)`, `IFERROR(0,${R})`, 'IFERROR(SEQUENCE(2),0)', 'IFERROR({1,2},0)', 'IFERROR(D!B1,0)', `IFNA(0,${R})`, 'IFNA(SEQUENCE(2),0)'],
  ifs: ['IFS(TRUE,1)', `IFS(TRUE,${R})`, 'SWITCH(1,1,2)', `SWITCH(1,1,${R})`],
  let: ['LET(x,1,x)', 'LET(x,D!B1,x)', `LET(x,${R},SUM(x))`, 'LET(x,{1,2},x)', 'LET(x,SEQUENCE(2),x)', `LET(x,${R},y,x,y)`, `LET(x,${R},INDEX(x,2))`, `LET(x,${R},ROWS(x))`, `LET(x,1,${R})`],
  lookup: ['VLOOKUP(5,D!L1:N6,{2,3},FALSE)', 'VLOOKUP(5,D!L1:N6,D!A12,FALSE)', 'HLOOKUP(5,D!L8:Q10,{1;2},FALSE)', 'XLOOKUP(5,D!L1:L6,D!M1:N6)', 'XLOOKUP(5,D!L1:N1,D!L2:N3)', 'XLOOKUP(5,D!L1:L6,{1;2;3;4;5;6})', 'XLOOKUP({5,6},D!L1:L6,D!M1:M6)', 'LOOKUP(5,D!L1:L6,D!M1:M6)', 'LOOKUP(5,D!L1:N6)', 'MATCH({5,6},D!L1:L6,0)'],
  rowcol: ['ROW(D!B1)', 'ROW(D!B1:C1)', 'COLUMN(D!B1)', 'COLUMN(D!B1:B3)', 'ROW(D!B:B)', 'COLUMN(D!1:1)', `ROW(INDEX(${M},0,1))`, 'ROW()', 'COLUMN()', 'SUM(ROW(D!B1))', 'SUM(COLUMN())', 'SUMPRODUCT(ROW(D!B1:B3))'],
  arrays: ['ABS({-1,-2})', 'SUM(ABS({-1,-2}))', 'SUM(COUNTIF(D!F1:F7,{"Чай","Чаи"}))', 'COUNTIF(D!F1:F7,{"Чай","Чаи"})', 'SUM({1,2}*2)', 'SUM(LEN({"a","bb"}))', 'ABS(ABS({-1,-2}))', 'LEN({"a","bb"})&""', 'ROWS({1;2}*2)', 'SUM(ABS(SEQUENCE(3)))', 'SUM(SEQUENCE(3)*2)', `SUM(${R}*{1;2;3})`, 'MAX({1,2})+1', 'SUM(TRANSPOSE(D!B1:B3))', 'SUM(MMULT({1,2},{3;4}))', 'SUM(INDEX({1,2;3,4},0,1))', 'INDEX({1,2;3,4},0,1)+0', '{1,2}&"a"', 'N({1,2})', 'T({"a","b"})'],
  shapes: ['D!B1:B1*2', 'D!B1:B1', `ISREF(${R})`, '(D!B1)', `SUM((${R}))`, `-${R}`, `+D!B1`, `SUM(+${R})`, `SUM(-${R})`, `D!B1:B3 D!B2:C2`],
}

export function buildProbes(): Probe[] {
  return Object.entries(GROUPS).flatMap(([group, fs]) => fs.map((f, i) => probe(`${group}#${i + 1}`, group, group, f)))
}
