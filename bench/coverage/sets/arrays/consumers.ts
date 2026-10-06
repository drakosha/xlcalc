/**
 * Functions that take arrays: aggregates, the lazy IF family, CHOOSE, criteria and lookups with array keys,
 * information and text functions lifted over ranges, INDEX with 0, ROW and COLUMN of several rows or columns.
 */
import type { Add } from './common.js'

export function aggregates(add: Add) {
  const g = (fn: string, ...fs: string[]) => add(`agg:${fn}`, fn, ...fs)
  g('SUM', 'SUM(D!B1:B8*{1;0;1;0;1;0;1;0})', 'SUM(IF(D!B1:B8>10,D!B1:B8))', 'SUM(IF(D!B1:B8>10,D!B1:B8,0))', 'SUM(ROUND(D!B1:B8/3,1))', 'SUM(LEN(D!F1:F7))', 'SUM(YEAR(D!H1:H7))')
  g('MAX', 'MAX(ABS(D!B1:B8))', 'MAX(IF(D!B1:B8<20,D!B1:B8))', 'MAX(D!F1:F7&"")', 'MAX(IF(D!B1:B8>20,ROW(D!B1:B8)))')
  g('MIN', 'MIN(IF(D!B1:B8>0,D!B1:B8))', 'MIN(IF(D!B1:B8>100,D!B1:B8))')
  g('AVERAGE', 'AVERAGE(IF(D!B1:B8>0,D!B1:B8))', 'AVERAGE(IF(D!B1:B8>100,D!B1:B8))', 'AVERAGE(D!B1:B8*1)')
  g('AVERAGEA', 'AVERAGEA(IF(D!B1:B8>10,D!B1:B8))', 'AVERAGEA({1,"a",TRUE})')
  g('COUNT', 'COUNT(IF(D!B1:B8>10,D!B1:B8))', 'COUNT(1/{1,0,2})')
  g('COUNTA', 'COUNTA(IF(D!B1:B8>10,D!B1:B8))', 'COUNTA(IF(D!B1:B8>10,D!B1:B8,""))', 'COUNTA(D!K1:K8&"")', 'COUNTA(IF(D!K1:K8="",D!K1:K8))')
  // a blank cell inside an array: kept by the functions that reshape, a value once lifted
  g('COUNTA', 'COUNTA(IF({1;1},D!K1:K2))', 'COUNTA(TRANSPOSE(D!K1:K2))', 'COUNTA(TAKE(D!K1:K8,2))', 'COUNTA(VSTACK(D!K1:K2))', 'COUNTA(SORT(D!D1:D8))', 'COUNTA(FILTER(D!D1:D8,{1;1;1;1;1;1;1;1}))')
  g('COUNTA', 'COUNTA(CHOOSECOLS(D!K1:K2,1))', 'COUNTA(INDEX(D!K1:K8,{1;2}))', 'COUNTA(INDEX(D!K1:K8,0,1))', 'COUNTA(UNIQUE(D!K1:K8))', 'COUNTA(TOCOL(D!K1:K3))', 'COUNTA(EXPAND(D!K1,2,1,1))', 'COUNTA(D!K1:K8*1)')
  g('ISBLANK', 'SUM(--ISNUMBER(IF({1;1},D!K1:K2)))', 'INDEX(IF({1;1},D!K1:K2),1)&"x"', 'INDEX(TRANSPOSE(D!K1:K2),1,1)&"x"', 'SUM(--ISBLANK(TRANSPOSE(D!K1:K2)))', 'SUM(--ISBLANK(SORT(D!D1:D8)))', 'SUM(--ISBLANK(INDEX(D!K1:K8,{1;2})))')
  g('MAXA', 'MAXA({TRUE,-1})', 'MAXA(IF(D!B1:B2>0,TRUE))')
  g('MINA', 'MINA({FALSE,1})')
  g('STDEVA', 'STDEVA({1,TRUE,3})')
  g('VARA', 'VARA({"a",2,4})')
  g('AVERAGEA', 'AVERAGEA({TRUE})', 'AVERAGEA({TRUE,"x"})')
  g('PRODUCT', 'PRODUCT(D!B1:B2*{1;2})')
  g('AND', 'AND(D!B1:B8>0)', 'AND({TRUE,1,"x"})', 'AND(IF(D!B1:B8>0,TRUE,"x"))')
  g('OR', 'OR(D!B1:B8>40)', 'OR({FALSE,0})')
  g('XOR', 'XOR({TRUE,TRUE,FALSE})')
  g('CONCAT', 'CONCAT({"a","b";"c","d"})', 'CONCAT(LEFT(D!F1:F7,1))', 'CONCAT(UPPER({"a","b"}))', 'CONCAT(D!K1:K3&"-")')
  g('TEXTJOIN', 'TEXTJOIN(",",TRUE,{1,2;3,4})', 'TEXTJOIN("-",FALSE,IF(D!B1:B8>10,D!B1:B8,""))', 'TEXTJOIN(",",TRUE,D!AS1:AS8&"")', 'TEXTJOIN(",",TRUE,TEXT(D!B1:B3,"0.0"))', 'TEXTJOIN({"-","+"},TRUE,{1,2,3})')
  g('SUMPRODUCT', 'SUMPRODUCT((D!B1:B8>10)*D!C1:C8)', 'SUMPRODUCT(--(D!F1:F7="чай"),D!G1:G7)', 'SUMPRODUCT(D!B1:B8>10)', 'SUMPRODUCT((D!L1:L6>2)*(D!N1:N6))', 'SUMPRODUCT(D!B1:B8,D!C1:C8*2)', 'SUMPRODUCT({1,2}*{3;4})', 'SUMPRODUCT({1,2},{3;4})', 'SUMPRODUCT((D!L1:L6>2)*ROW(D!L1:L6))')
  g('LARGE', 'LARGE(D!B1:B8*2,2)', 'LARGE({3,1,2},{1,2})', 'SUM(LARGE(D!B1:B8,{1,2,3}))')
  g('SMALL', 'SMALL(IF(D!B1:B8>5,D!B1:B8),2)')
  g('MEDIAN', 'MEDIAN(D!B1:B8-D!C1:C8)')
  g('STDEV.S', 'STDEV.S(IF(D!B1:B8>0,D!B1:B8))')
  g('VAR.P', 'VAR.P({1,2,3,4}*2)')
  g('MODE.SNGL', 'MODE.SNGL(D!AS1:AS8*1)', 'MODE.SNGL(IFERROR(D!AS1:AS8*1,""))')
  g('GCD', 'GCD({12,18,24})')
  g('CORREL', 'CORREL(D!B1:B8*2,D!C1:C8)')
  g('NPV', 'NPV(0.1,{-100,50,60})')
  g('IRR', 'IRR({-100,50,60})')
}

export function lazy(add: Add) {
  const g = (fn: string, ...fs: string[]) => add(`lazy:${fn}`, fn, ...fs)
  g('IF', 'SUM(IF({TRUE,FALSE,TRUE},{1,2,3},{10,20,30}))', 'SUM(IF(D!B1:B8>10,1,0))', 'SUM(IF(D!B1:B8>10,{1;2;3;4;5;6;7;8}))', 'INDEX(IF({1,0},D!B1:B2,D!C1:C2),2,2)', 'SUM(IF(D!E1:E8>3,D!E1:E8))')
  g('IF', 'ROWS(IF(TRUE,D!B1:B3))', 'INDEX(IF(D!B1:B2>5,"y","n"),1)', 'SUM(IF(D!B1:B3>5,D!C1:C2,100))', 'SUM(IF({1;0},{1,2},{3,4,5}))', 'SUM(--ISBLANK(IF(TRUE,D!K1:K8)))', 'SUM(--ISBLANK(IF({1;1;1;1;1;1;1;1},D!K1:K8)))', 'IF(D!B1:B2>5,"y")')
  g('IFERROR', 'SUM(IFERROR(IF(D!E1:E8>3,D!E1:E8),0))', 'SUM(IFERROR(D!E1:E8,100))', 'SUM(IFERROR(1/{1,0,2},100))', 'SUM(IFERROR(1/{1,0,2},{10,20,30}))', 'ROWS(IFERROR(D!B1:B3,0))', 'ROWS(IFERROR(1,D!B1:B3))')
  g('IFNA', 'SUM(IFNA({1,#N/A,3},10))', 'SUM(IFNA(MATCH({2,4,8},D!L1:L6,0),0))')
  g('IFS', 'SUM(IFS(D!B1:B8>20,1,TRUE,0))', 'INDEX(IFS({TRUE,FALSE},"a",{FALSE,TRUE},"b"),1,2)')
  g('SWITCH', 'SUM(SWITCH(D!L1:L6,1,10,2,20,0))', 'INDEX(SWITCH({1,2},1,"one","other"),1,2)')
  g('NOT', 'SUM(--NOT(D!B1:B8>10))')
  g('CHOOSE', 'SUM(CHOOSE(2,D!B1:B8,D!C1:C8))', 'SUM(CHOOSE({1,2},10,20))', 'SUMIF(CHOOSE(1,D!L1:L6,D!N1:N6),">2")', 'ROWS(CHOOSE(1,D!B1:B8,1))', 'INDEX(CHOOSE({1,2},D!B1:B2,D!C1:C2),2,2)')
}

export function keyed(add: Add) {
  const g = (fn: string, ...fs: string[]) => add(`key:${fn}`, fn, ...fs)
  g('COUNTIF', 'SUM(COUNTIF(D!F1:F7,{"аренда","чай"}))', 'INDEX(COUNTIF(D!B1:B8,D!B1:B8),3)', 'SUM(COUNTIF(D!AS1:AS8,D!AS1:AS8))', 'SUM(1/COUNTIF(D!AS1:AS8,D!AS1:AS8))')
  g('SUMIF', 'SUM(SUMIF(D!L1:L6,{1,3},D!N1:N6))')
  g('SUMIFS', 'SUM(SUMIFS(D!N1:N6,D!L1:L6,{">2","<2"}))')
  g('COUNTIFS', 'COUNTIFS(D!L1:L6,">"&{1,5})', 'SUM(COUNTIFS(D!L1:L6,">"&{1,5}))', 'SUM(COUNTIFS(D!L1:L6,{">1",">5"},D!N1:N6,"<100"))')
  g('AVERAGEIF', 'SUM(AVERAGEIF(D!L1:L6,{">2","<3"},D!N1:N6))')
  g('VLOOKUP', 'SUM(VLOOKUP({1,3},D!L1:N6,3,FALSE))', 'INDEX(VLOOKUP(2,D!L1:N6,{2,3},FALSE),1,2)', 'VLOOKUP(2,D!L1:N6,{2,3},FALSE)', 'SUM(VLOOKUP(D!L1:L3,D!L1:N6,3))')
  g('HLOOKUP', 'SUM(HLOOKUP({2,3},D!L8:Q10,3,FALSE))')
  g('MATCH', 'SUM(MATCH({2,8},D!L1:L6,0))', 'MATCH(1,(D!L1:L6>2)*(D!N1:N6>40),0)', 'MATCH(1,{0,0,1},0)', 'MATCH("b*",{"a","bc"},0)')
  g('XLOOKUP', 'ROWS(XLOOKUP(9,D!L1:L6,D!M1:M6,D!B1:B2))', 'SUM(XLOOKUP({2,8},D!L1:L6,D!N1:N6))', 'INDEX(XLOOKUP(2,D!L1:L6,D!M1:N6),1,2)', 'COLUMNS(XLOOKUP(2,D!L1:L6,D!M1:N6))', 'XLOOKUP(2,D!L1:L6,D!M1:N6)', 'INDEX(XLOOKUP(2,D!L8:Q8,D!L9:Q10),2,1)', 'XLOOKUP(9,D!L1:L6,D!M1:N6)', 'INDEX(XLOOKUP(9,D!L1:L6,D!M1:N6,{"x","y"}),1,2)')
  g('XMATCH', 'SUM(XMATCH({3,5},D!L1:L6))', 'XMATCH(3,{1,3,5})')
  g('LOOKUP', 'SUM(LOOKUP({2,6},D!L1:L6,D!N1:N6))', 'LOOKUP(2,{1,2,3},{"a","b","c"})')
}

export function lifted(add: Add) {
  const g = (fn: string, ...fs: string[]) => add(`lifted:${fn}`, fn, ...fs)
  g('ISBLANK', 'SUM(--ISBLANK(D!K1:K8))', 'SUM(--ISBLANK(D!K1:K8&""))', 'SUM(--ISBLANK(D!D1:D8))')
  g('ISNUMBER', 'SUM(--ISNUMBER(D!D1:D8))', 'SUM(--ISNUMBER(D!D1:D8*1))')
  g('ISTEXT', 'SUM(--ISTEXT(D!D1:D8))')
  g('ISERROR', 'SUM(--ISERROR(D!E1:E8))')
  g('TYPE', 'TYPE(D!B1:B2)', 'TYPE(D!B1:B2*1)', 'TYPE(D!B1)', 'TYPE(D!K1:K2)')
  g('N', 'N(D!B1:B3)', 'SUM(N(D!B1:B3))', 'N({1,2})')
  g('T', 'T(D!F1:F2)', 'T({"a","b"})')
  g('EXACT', 'SUM(--EXACT(D!F1:F7,"аренда"))')
  g('MID', 'INDEX(MID("abcdef",{1,3,5},2),1,3)')
  g('DATE', 'INDEX(DATE(2025,{1,2,3},1),1,3)')
  g('EOMONTH', 'SUM(EOMONTH(D!H1:H2,0))', 'SUM(EOMONTH(D!H1:H2+0,0))', 'SUM(EOMONTH({45658,45689},0))')
  g('EDATE', 'SUM(EDATE({45658,45689},1))', 'SUM(EDATE(D!H1:H2,1))')
  g('TEXT', 'INDEX(TEXT(5,{"0","0.0"}),1,2)', 'TEXT(D!B1:B2,"0")')
  g('ROUND', 'INDEX(ROUND(D!B1:B8,{0,1}),7,2)')
  g('VALUETOTEXT', 'VALUETOTEXT(D!B1:B2)', 'INDEX(VALUETOTEXT({1,"a"},1),1,2)')
  g('ARRAYTOTEXT', 'ARRAYTOTEXT({1,2;3,4},1)', 'ARRAYTOTEXT(D!B1:B2*2)')
  g('ERROR.TYPE', 'SUM(IFERROR(ERROR.TYPE(D!E1:E3),0))')
  g('BESSELJ', 'BESSELJ(D!B1:B2,1)', 'SUM(BESSELJ(D!B1:B2+0,1))', 'SUM(BESSELJ({1,2},1))')
  g('DEC2BIN', 'DEC2BIN(D!B1:B2)', 'INDEX(DEC2BIN({4,8}),1,2)')
  g('IMABS', 'IMABS(D!B1:B2)', 'SUM(IMABS({"3+4i","1"}))')
}

export function reference(add: Add) {
  const g = (fn: string, ...fs: string[]) => add(`ref:${fn}`, fn, ...fs)
  g('INDEX', 'SUM(INDEX(D!L1:N6,0,3))', 'SUM(INDEX(D!L1:N6,2,0))', 'ROWS(INDEX(D!L1:N6,0,0))', 'COLUMNS(INDEX(D!L1:N6,0,0))', 'SUM(INDEX(D!L1:N6,0,0))', 'INDEX(D!L1:N6,0,2)', 'COUNTA(INDEX(D!L1:N6,0,2))')
  g('INDEX', 'INDEX(INDEX(D!L1:N6,0,2),3)', 'SUM(INDEX(D!L1:N6,{1;2},{1,3}))', 'INDEX(D!L1:N6,{1;2},2)', 'SUMIF(INDEX(D!L1:N6,0,1),">2",INDEX(D!L1:N6,0,3))', 'SUM(INDEX({1,2;3,4},0,2))', 'SUM(INDEX({1,2;3,4},2,0))')
  g('INDEX', 'SUM(INDEX(D!B1:B8,D!AH2:AH4))', 'INDEX(D!K1:K8*1,1)', 'INDEX({1,2;3,4},0,3)', 'SUM(INDEX(D!B:B,0,1))', 'ROWS(INDEX(D!B:B,0,1))', 'SUM(INDEX(D!L1:N6*1,0,1))', 'SUM(INDEX(D!L1:L6,0))', 'SUM(INDEX(D!L8:Q8,0))', 'INDEX(D!L1:N6,0)', 'SUM(INDEX(D!L1:N6,{1,2,3},1))')
  g('ROW', 'SUM(ROW(D!C1:C8))', 'ROWS(ROW(D!L1:N6))', 'COLUMNS(ROW(D!L1:N6))', 'INDEX(ROW(D!A3:A9),2)', 'SUM(ROW(D!1:3))', 'ROW(D!C1:C8)')
  g('COLUMN', 'SUM(COLUMN(D!L1:N6))', 'COLUMNS(COLUMN(D!L1:N6))', 'COLUMNS(COLUMN(D!1:1))', 'INDEX(COLUMN(D!B1:E1),1,4)')
  g('ROWS', 'ROWS({1;2;3})', 'ROWS(D!B1:B8*1)', 'ROWS(5)')
  g('COLUMNS', 'COLUMNS({1;2;3})', 'COLUMNS({1,2}*{1;2;3})')
}
