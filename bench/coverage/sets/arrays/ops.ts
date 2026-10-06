/** Operators over arrays and ranges, array constants, and what a formula with an array result gives. */
import type { Add } from './common.js'

export function operators(add: Add) {
  const g = (...fs: string[]) => add('ops', 'array-ops', ...fs)
  // each operator, element by element
  g('SUM(D!B1:B8*2)', 'SUM(D!B1:B8*D!C1:C8)', 'SUM(D!B1:B8-D!C1:C8)', 'SUM(D!B1:B8/D!C1:C8)', 'SUM(D!B1:B8^2)', 'SUM(-D!B1:B8)', 'SUM(D!B1:B8%)', 'SUM(+D!B1:B8)', 'INDEX(D!B1:B8%,1)')
  g('SUM(D!B1:B8*D!B1)', 'SUM(2^{1,2,3})', 'SUM({2,3}^{2;3})', 'INDEX({0}^{0},1)', 'INDEX(1/{1,0,2},1,2)', 'SUM(1E+307*{100,1})', 'INDEX(-{"1","a"},1,1)', 'INDEX(-{"1","a"},1,2)')
  // sizes: the larger rows and columns, a single row or column repeated, #N/A outside
  g('SUM({1,2,3}+{10;20})', 'ROWS({1,2,3}+{10;20})', 'COLUMNS({1,2,3}+{10;20})', 'SUM({1,2,3}+{1,2})', 'INDEX({1,2,3}+{1,2},1,3)', 'INDEX({1,2,3}+{1,2},1,2)', 'COLUMNS({1,2,3}+{1,2})')
  g('INDEX({1,2;3,4}+{10,20,30},2,3)', 'INDEX({1,2;3,4}+{10,20,30},2,2)', 'INDEX({1;2}*{1;2;3},3)', 'ROWS(D!B1:B3*D!C1:C2)', 'INDEX(D!B1:B3*D!C1:C2,3)', 'INDEX(D!B1:B3*D!C1:C2,2)')
  g('SUM(D!P1:Q2*D!B1:B2)', 'SUM(D!B1:B2*D!P1:Q1)', 'ROWS(D!B1:B2*D!P1:Q1)', 'COLUMNS(D!B1:B2*D!P1:Q1)', 'INDEX(D!P4:R5*{10;100},2,3)', 'INDEX(D!P4:R5+{1,2},1,3)')
  // values inside: numeric text, text, booleans, blanks, errors
  g('INDEX(D!D1:D8*1,2)', 'INDEX(D!D1:D8*1,3)', 'INDEX(D!D1:D8*1,4)', 'INDEX(D!D1:D8*1,5)', 'SUM(IFERROR(D!D1:D8*1,0))', 'SUM(D!D1:D8*1)')
  g('SUM(D!E1:E8*1)', 'INDEX(D!E1:E8*1,2)', 'COUNT(D!E1:E8*1)', 'INDEX(D!K1:K8+D!K1:K8,1)', 'INDEX(D!K1:K8&"",1)', 'LEN(CONCAT(D!K1:K8&"x"))')
  // comparisons
  g('SUM((D!B1:B8>10)*1)', 'SUM(--(D!F1:F7="аренда"))', 'SUM(--(D!D1:D8=""))', 'SUM(--(D!D1:D8=0))', 'SUM(--(D!D1:D8>2))', 'SUM(--(D!D1:D8=FALSE))', 'SUM(--(D!E1:E8>3))', 'SUM(--({"a","B","c"}<"b"))')
  // the & operator
  g('INDEX(D!F1:F7&"!",2)', 'INDEX("<"&D!B1:B8&">",7)', 'CONCAT(D!B1:B3&",")', 'INDEX(D!D1:D8&"",4)', 'INDEX({1,2}&{"a";"b"},2,1)')
  // whole columns: 1,048,576 rows, the blank ones too
  g('SUM(D!B:B*1)', 'SUM(D!B:B+1)', 'COUNT(D!B:B*1)', 'ROWS(D!B:B*1)', 'SUM(D!B:B*D!C:C)', 'SUMPRODUCT((D!B:B>10)*D!C:C)', 'MAX(D!B:B*2)', 'MIN(D!B:B*2)', 'MIN(D!B:B+100)', 'COUNTA(D!B:B*1)')
  g('COLUMNS(D!1:1*0)', 'MATCH(1,(D!B:B>20)*(D!C:C>20),0)', 'MATCH(0,(D!B:B>20)*1,0)', 'XMATCH(1,(D!B:B="")*1)', 'SUM(--(D!B:B=""))', 'AVERAGE(D!B:B*1)')
}

export function constants(add: Add) {
  const g = (...fs: string[]) => add('const', 'array-constants', ...fs)
  g('SUM({1,2;3,4})', 'ROWS({1,2;3,4})', 'COLUMNS({1,2,3})', 'INDEX({1,2;3,4},2,1)', 'INDEX({"a","b"},2)', 'INDEX({TRUE,FALSE},2)', 'INDEX({1,#N/A},2)', 'INDEX({-1,-2.5},2)')
  g('INDEX({1E+3,1.5E-2},1)', 'INDEX({1E+3,1.5E-2},2)', 'COUNTA({1,"a",TRUE,#N/A})', 'COUNT({1,"2",TRUE})', 'SUM({1,"2",TRUE})', 'AVERAGE({1,2,"x"})', 'SUM({5})', '{5}', '{1,2}', '{"a";"b"}')
  g('INDEX({"a""b"},1)', 'INDEX({1;2;3},3)', 'TYPE({1,2})', 'TYPE({1})', 'INDEX({1,2},1,1)+INDEX({3;4},2,1)', 'INDEX({#DIV/0!,#VALUE!},1,2)', 'ISERROR(INDEX({#N/A},1))', 'INDEX({true,false},1)')
  g('MAX({"a",TRUE})', 'MIN({-1,"-5",FALSE})', 'PRODUCT({2,3,4})', 'COUNTA({""})', 'LEN(INDEX({""},1))', 'INDEX({0.1,-0},1,2)', 'SUM({1,2,3}*{1,2,3})', 'INDEX({1,2;3,4},0,1)')
}

/** The result of a formula is an array: more than one value is #SPILL! (every neighbour is taken), one value is that value. */
export function results(add: Add) {
  const g = (...fs: string[]) => add('result', 'array-result', ...fs)
  g('D!B1:B3', 'D!B1:B1', 'D!B1:B8*1', '{1;2}*{3}', '{7}*2', 'D!B1:B2+0', 'IF(TRUE,D!B1:B2)', 'IF(TRUE,D!B1)', 'IF({TRUE},1,2)', 'IF({TRUE,FALSE},1,2)')
  g('INDEX(D!L1:N6,0,1)', 'INDEX(D!L1:N1,0,1)', 'INDEX(D!L1:L6,0,1)', 'ROW(D!A1:A3)', 'ROW(D!A3:C3)', 'COLUMN(D!A1:A3)', 'COLUMN(D!A1:C1)', 'CHOOSE(2,D!B1:B2,D!C1)', 'CHOOSE(1,D!B1:B2,D!C1)')
  g('D!K1:K2', 'D!K1', 'D!E3:E4', 'D!A1:A1', 'D!B:B', 'D!1:1', 'IF(D!B1>0,D!B1:B2,0)', 'IF(D!B1>100,D!B1:B2,0)', 'ABS(D!B7:B7)', 'ABS(D!B7:B8)')
}
