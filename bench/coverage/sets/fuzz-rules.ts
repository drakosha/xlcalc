/**
 * Excel rules the random check found (`fuzz.ts`), measured case by case so each stays pinned: text read as
 * VALUE reads it in numeric arguments, the order in which Excel reads arguments and checks them (ToolPak
 * functions, values before references, one-cell errors), collinear variables in LINEST, the database field,
 * empty delimiters, odd coupon periods before 1900. Computed as dynamic array formulas (Excel 365). Answers:
 * `test/excel/cases/fuzz-rules.json`.
 */
import type { Probe } from '../probes.js'
import { collector } from './arrays/common.js'

export const DYNAMIC = true

/** Cases by function: `fn` is the function checked. */
const CASES: Record<string, string[]> = {
  // text as a number, as VALUE reads it: dates, times, %, currency, groups
  ABS: ['ABS("2025-01-31")', 'ABS("$1,000")', 'ABS("(5)")', 'ABS("1 1/2")'],
  LN: ['LN("50%")'],
  SIGN: ['SIGN("$1,000")', 'SIGN("-1,000.5")'],
  ROUNDDOWN: ['ROUNDDOWN(3.2,"$1,000")'],
  'NORM.S.DIST': ['NORM.S.DIST("2025-01-31 12:00",TRUE)', 'NORM.S.DIST("6:35 AM",TRUE)'],
  TEXT: ['TEXT("2025-01-31 12:00","0")', 'TEXT("50%","0.00")', 'TEXT("abc","0")', 'TEXT("$1,000","#,##0")'],
  DOLLAR: ['DOLLAR("1/2/2025",1)'],
  BAHTTEXT: ['BAHTTEXT("50%")'],
  RIGHTB: ['RIGHTB("2025-01-31","2025-01-31 12:00")'],
  EDATE: ['EDATE(40558,"50%")', 'EDATE(40558,"1E1")'],
  SUM: ['SUM("50%",1)', 'SUM({1,#DIV/0!},NA())', 'SUM(D!E1:E8,NA())', 'SUM(NA(),D!E1:E8)', 'SUM(D!E1:E8,"x")'],
  // operators read text the same way
  '+': ['"31-Jan-2025"-"2025-01-31"', '"50%"*2', '"$1,000"+1', '"18:30"*24', '" 2 "+1', '"1,5"+0'],
  // ToolPak securities read the basis first
  DISC: ['DISC(NA(),48214,97.975,100,7)', 'DISC(39264,48214,NA(),100,9)', 'DISC(48214,39264,97.975,100,NA())', 'DISC(48214,39264,97.975,100,"x")'],
  INTRATE: ['INTRATE(NA(),39583,1000000,1014420,7)', 'INTRATE(39583,39493,1000000,1014420,NA())'],
  PRICEDISC: ['PRICEDISC(39494,"x y",0.0525,100,7)', 'PRICEDISC(39494,39508,0.0525,TRUE,7)', 'PRICEDISC(NA(),39508,0.0525,100,7)', 'PRICEDISC(39494,39508,0.0525,100,NA())', 'PRICEDISC(39508,39494,NA(),100,2)'],
  RECEIVED: ['RECEIVED(NA(),39583,1000000,0.0575,7)', 'RECEIVED("abc",39583,NA(),0.5,20)'],
  YIELDDISC: ['YIELDDISC("abc",39508,15,100,7)', 'YIELDDISC("abc",39508,15,100,0)'],
  PRICEMAT: ['PRICEMAT(NA(),39551,39397,0.061,0.061,7)', 'PRICEMAT(39551,39493,39397,0.061,0.061,9)'],
  YIELDMAT: ['YIELDMAT(39522,39755,39395,NA(),100.0123,7)'],
  ACCRINTM: ['ACCRINTM(NA(),39614,0.1,1000,7)'],
  PRICE: ['PRICE(NA(),42689,0.0575,0.065,100,2,7)', 'PRICE(39493,42689,0.0575,0.065,100,3,NA())', 'PRICE(42689,39493,0.0575,0.065,100,3,0)'],
  YIELD: ['YIELD(39493,42689,NA(),95.04287,100,2,7)', 'YIELD(39493,42689,0.0575,95.04287,100,NA(),9)'],
  YEARFRAC: ['YEARFRAC("x",40909,9)', 'YEARFRAC(40909,"x",NA())', 'YEARFRAC(NA(),41120,7)'],
  // coupons check the dates, then the frequency, then the basis, as they read them
  COUPDAYBS: ['COUPDAYBS(40862,40568,2,NA())', 'COUPDAYBS(40862,40568,2,"x")', 'COUPDAYBS(40568,40862,3,NA())', 'COUPDAYBS(NA(),40862,3,1)'],
  COUPNUM: ['COUPNUM(40862,40568,2,NA())', 'COUPNUM(40568,40862,3,NA())'],
  COUPNCD: ['COUPNCD(40862,40568,NA(),1)', 'COUPNCD(40568,40862,"x",9)'],
  COUPPCD: ['COUPPCD(45688,20,"x y")', 'COUPPCD(DATE(2011,1,25),DATE(2011,11,15),0,#VALUE!)'],
  DURATION: ['DURATION(43282,54058,0.08,0.09,3,NA())', 'DURATION(54058,43282,0.08,0.09,2,NA())'],
  MDURATION: ['MDURATION(43282,54058,0.08,0.09,3,NA())'],
  ACCRINT: ['ACCRINT(39508,39691,39569,0.1,1000,3,NA())', 'ACCRINT(39569,39691,39508,0.1,1000,2,NA())'],
  ODDFPRICE: ['ODDFPRICE(39763,44256,39736,39873,0.0785,0.0625,100,3,NA())', 'ODDFPRICE(-1,59,FALSE,39873,0.0785,"x",100,2,1)', 'ODDFPRICE(400,44256,60,39873,0.0785,0.0625,100,2,1)', 'ODDFPRICE(400,44256,61,39873,0.0785,0.0625,100,2,1)', 'ODDFPRICE(400,44256,59,39873,0.0785,0.0625,100,4,1)'],
  ODDFYIELD: ['ODDFYIELD(39763,44256,39736,39873,0.0575,84.5,100,3,NA())'],
  // an odd last period whose quasi-coupon period on the maturity's schedule begins before 1900 is #NUM!
  ODDLPRICE: [
    'ODDLPRICE(39485,39614,39370,0.0375,"3",100,-3.5,#N/A)', 'ODDLPRICE(39485,39614,164,0.0375,0.0405,100,2,0)', 'ODDLPRICE(39485,39614,168,0.0375,0.0405,100,2,0)',
    'ODDLPRICE(39485,39614,72,0.0375,0.0405,100,4,0)', 'ODDLPRICE(39485,39614,75,0.0375,0.0405,100,4,0)', 'ODDLPRICE(39485,39614,120,0.0375,0.0405,100,1,0)', 'ODDLPRICE(39485,39614,182,0.0375,0.0405,100,1,0)',
  ],
  ODDLYIELD: ['ODDLYIELD(39558,39614,39440,0.0375,99.875,100,3,NA())', 'ODDLYIELD(39558,39614,100,0.0375,99.875,100,2,0)', 'ODDLYIELD(39558,39614,184,0.0375,99.875,100,2,0)'],
  AMORDEGRC: ['AMORDEGRC(-1,39679,39813,300,1,0.15,NA())', 'AMORDEGRC(2400,39679,39813,300,-1,0.15,NA())', 'AMORDEGRC(2400,39679,39813,3000,1,0.15,NA())', 'AMORDEGRC(-1,39679,39813,NA(),1,0.15,1)', 'AMORDEGRC(2400,39813,39679,300,1,0.15,NA())'],
  AMORLINC: ['AMORLINC(2400,39679,39813,300,1,-0.15,NA())', 'AMORLINC(2400,39813,39679,300,1,0.15,"x")'],
  // the radix functions read their places first, DELTA and GESTEP their second number, COMPLEX its suffix
  DEC2BIN: ['DEC2BIN(NA(),"x")', 'DEC2BIN(NA(),-1)', 'DEC2BIN("x",-1)', 'DEC2BIN(NA(),TRUE)', 'DEC2BIN(1,NA())'],
  DEC2HEX: ['DEC2HEX(NA(),"x")'],
  HEX2BIN: ['HEX2BIN(NA(),"x")', 'HEX2BIN(NA(),11)'],
  OCT2HEX: ['OCT2HEX(NA(),"x")'],
  DELTA: ['DELTA(NA(),"x")', 'DELTA("x",NA())'],
  GESTEP: ['GESTEP(NA(),"x")', 'GESTEP("x",NA())'],
  COMPLEX: ['COMPLEX(3,NA(),"x")', 'COMPLEX(NA(),4,"x")', 'COMPLEX(NA(),4)', 'COMPLEX(1,1,NA())', 'COMPLEX(1,1,"")'],
  FVSCHEDULE: ['FVSCHEDULE(#VALUE!,D!A8)', 'FVSCHEDULE(1,D!A8)', 'FVSCHEDULE(NA(),{1,"x"})', 'FVSCHEDULE(TRUE,D!A8)'],
  ISPMT: ['ISPMT(0.1,1.5,4,1000)', 'ISPMT(0.1,1,4.5,1000)', 'ISPMT(0.1,-1.5,4,1000)'],
  // values before references and arrays; GCD keeps the order
  LARGE: ['LARGE(D!E1:E8,NA())', 'LARGE({1,#N/A},"x")'],
  SMALL: ['SMALL(D!E1:E8,"x")'],
  QUARTILE: ['QUARTILE(D!E1:E8,NA())'],
  'PERCENTILE.EXC': ['PERCENTILE.EXC(D!A7,#NAME?)', 'PERCENTILE.EXC(D!E1:E8,NA())'],
  'PERCENTILE.INC': ['PERCENTILE.INC(D!E1:E8,"x")'],
  'PERCENTRANK.INC': ['PERCENTRANK.INC(D!E1:E8,"x")', 'PERCENTRANK.INC(D!E1:E8,5,"x")'],
  PERCENTRANK: ['PERCENTRANK(D!A8,#VALUE!)'],
  TRIMMEAN: ['TRIMMEAN(D!E1:E8,"x")'],
  PROB: ['PROB(D!AH1:AH4,{0.2,#N/A,0.1,0.4},"x")'],
  MIRR: ['MIRR(D!I13:I17,"x",0.1)'],
  ZTEST: ['ZTEST(D!E1:E8,NA())'],
  MAX: ['MAX(D!E1:E8,"x")'],
  AVERAGE: ['AVERAGE({1,#N/A},1/0)'],
  MAXA: ['MAXA({1,#N/A},1/0)'],
  GCD: ['GCD({4,#N/A},1/0)', 'GCD({4,"a"},1/0)', 'GCD(D!B1:C8,NA())'],
  CONCAT: ['CONCAT({"a",#DIV/0!},NA())', 'CONCAT(NA(),{"a",#DIV/0!})'],
  AND: ['AND({TRUE,#DIV/0!},NA())'],
  TEXTJOIN: ['TEXTJOIN(",",TRUE,{"a",#DIV/0!},NA())'],
  // pairs: a one-cell error, then the sizes, then errors inside
  STEYX: ['STEYX(D!C1:C8,{1,#NUM!})', 'STEYX(D!C1:C8,D!B1:B8/0)'],
  'COVARIANCE.P': ['COVARIANCE.P(D!B1:B8,{#VALUE!,1})'],
  PEARSON: ['PEARSON(D!B1:B8,{#VALUE!,1})'],
  SLOPE: ['SLOPE(D!B1:B8,{#VALUE!,1})'],
  CORREL: ['CORREL(D!A7,D!B1:B8)'],
  CHITEST: ['CHITEST({1,#VALUE!},D!AM1:AN3)', 'CHITEST(D!AK1:AL3,{#VALUE!,1;1,1;1,1})', 'CHITEST({1,#N/A},{1,1,1})'],
  TTEST: ['TTEST(D!B1:B8,D!A8,0.5,1)', 'TTEST(D!B1:B8,{1,#N/A},0.5,1)', 'TTEST(D!B1:B8,D!C1:C5,0.5,1)', 'TTEST(D!B1:B8,D!A8,1,2)'],
  // the database field: a cell holding an error is #VALUE!, an error value itself; a database without records #VALUE!
  DCOUNTA: ['DCOUNTA(D!U1:Y7,D!A7,D!AA1:AB2)', 'DCOUNTA(D!U1:Y7,NA(),D!AA1:AB2)', 'DCOUNTA(D!AY1:AY5,1,D!AZ1:AZ5)'],
  DVAR: ['DVAR(D!B3,1,D!F1:F7)', 'DVAR(D!U1:Y7,TRUE,D!F1:F7)'],
  DSUM: ['DSUM(D!B3,1,D!AA1:AB2)', 'DSUM(D!U1:Y7,D!X1,D!AA1:AB2)', 'DSUM(D!U1:Y7,D!U1:V1,D!AA1:AB2)', 'DSUM(D!U1:Y1,4,D!AA1:AB2)'],
  DCOUNT: ['DCOUNT(D!B3:B4,1,D!AA1:AB2)'],
  // an empty delimiter in TEXTSPLIT is #VALUE!
  TEXTSPLIT: ['TEXTSPLIT("a,b",{",",""})', 'TEXTSPLIT("a,b",",",{";",""})', 'TEXTSPLIT("a,b",",",D!K1)', 'TEXTSPLIT("a,b;c",D!K1:K8,";")'],
  // PERCENTOF: an error given as a value is the answer, one in a range or array #NUM!
  PERCENTOF: ['PERCENTOF(2,NA())', 'PERCENTOF(NA(),2)', 'PERCENTOF({1,#N/A},3)', 'PERCENTOF(1,D!E1:E8)'],
  WEEKNUM: ['WEEKNUM(40977,1000)', 'DAY(WEEKNUM(40977,1E+3)<>2.5)'],
  // batch 2: complex functions read number text as VALUE does unless a sign follows the first character
  IMCONJUGATE: ['IMCONJUGATE("$8.00")'],
  IMLOG10: ['IMLOG10("$1,000")'],
  IMEXP: ['IMEXP("18:30")'],
  IMABS: ['IMABS("50%")'],
  IMSUM: ['IMSUM("2025-01-31",1)', 'IMSUM("18:30",1)'],
  IMREAL: ['IMREAL("(5)")', 'IMREAL("1,000")', 'IMREAL("1/2/2025")', 'IMREAL("31-Jan-2025")', 'IMREAL("6:35 AM")', 'IMREAL("1 1/2")', 'IMREAL("-$5")', 'IMREAL("TRUE")', 'IMREAL("25:00")', 'IMREAL("2025-01-31 12:00")'],
  // XIRR: a boolean or other text among the dates is #VALUE! before a blank date's #NUM!
  XIRR: ['XIRR(D!AY1:AY5,D!A1:A5,D!A2)', 'XIRR(D!J1:J5,D!A1:A5)', 'XIRR({-1,2},{1,TRUE})', 'XIRR({1,2},{1,TRUE})', 'XIRR(D!J1:J2,D!A1:A2)', 'XIRR(D!I1:I5,D!A1:A5)'],
  // GCD, LCM, MULTINOMIAL read every argument in order, inside ranges and arrays too, and check signs last
  MULTINOMIAL: ['MULTINOMIAL(2,3,{2,"a";TRUE,#N/A})', 'MULTINOMIAL(NA(),{"a"})'],
  LCM: ['LCM(D!A3,NA())', 'LCM(-0.5,2)'],
  // whole numbers near an integer (Excel takes some as that integer)
  INT: ['INT(FTEST(D!C1:C8,D!C1:C8))'],
  LEFT: ['LEFT("abc",0.9999999999999997)'],
  INDEX: ['INDEX({1,2,3},1,1.9999999999999997)'],
  MID: ['MID("abcdef",1.9999999999999997,2)'],
}

/** LINEST and its kin with collinear variables (left out with coefficient 0) and a single observation. */
function fits(): Record<string, string[]> {
  const at = (f: string, cells: [number, number][]) => cells.map(([r, c]) => `INDEX(${f},${r},${c})`)
  const row = (n: number): [number, number][] => Array.from({ length: n }, (_, j) => [1, j + 1])
  const stats: [number, number][] = [[2, 1], [2, 2], [2, 3], [3, 1], [3, 2], [4, 1], [4, 2], [5, 1], [5, 2]]
  return {
    LINEST: [
      ...at('LINEST({1;2;3;5},{1,2;2,4;3,6;4,8},TRUE,TRUE)', [...row(3), ...stats]),
      ...at('LINEST({1;2;3},{1,1;2,2;3,3})', row(3)),
      ...at('LINEST({1;2;3},{1,1;2,2;3,3},FALSE,TRUE)', [...row(2), [4, 2]]),
      ...at('LINEST({1;2;4;5},{1,0.5;2,1;3,1.5;5,2.5})', row(3)),
      ...at('LINEST({1;2;4},{1,2;2,4;3,6},FALSE)', row(2)),
      ...at('LINEST({1;2;4;5},{10,0;11,1;12,2;13,3})', row(3)),
      ...at('LINEST({1;2;4;7},{1,1,2;2,1,3;3,2,5;4,2,6},TRUE,TRUE)', [...row(4), [4, 1], [4, 2]]),
      ...at('LINEST({1;2;4;7},{1,1,2;2,1,3;3,2,5;4,2,6},FALSE)', row(3)),
      ...at('LINEST({1;2;3;5},{1,1,2;2,1,4;3,1,6;4,1,8},TRUE,TRUE)', [...row(4), [2, 3], [2, 4], [4, 1], [4, 2]]),
      ...at('LINEST({1,2,3},{1,2,3;2,4,6})', row(3)),
      ...at('LINEST({1;2;3},{0;0;0})', row(2)),
      ...at('LINEST({1;2;3},{0;0;0},FALSE,TRUE)', [[1, 1], [4, 2]]),
      ...at('LINEST({2;2;2},{1;2;3},TRUE,TRUE)', [[3, 1], [3, 2], [4, 1]]),
      ...at('LINEST(5,{1,2,3},TRUE,TRUE)', [...row(4), [2, 1], [2, 4], [3, 1], [3, 2], [4, 2]]),
      ...at('LINEST(5,{1,2,3},FALSE)', row(4)),
      ...at('LINEST(5,D!BC1:BC6)', [[1, 1], [1, 7]]),
      ...at('LINEST(D!B1,D!C1,TRUE,TRUE)', [[1, 1], [1, 2], [2, 1], [2, 2], [3, 1], [3, 2], [4, 1], [4, 2], [5, 1]]),
      'COLUMNS(LINEST(5,D!BC1:BC6))', 'COLUMNS(LINEST(5,{1,2,3}))',
      // sizes before values
      'INDEX(LINEST(D!L1:N6,D!BC1:BC6),1,1)', 'INDEX(LINEST(D!B1:B8,D!L1:N6),1,1)', 'INDEX(LINEST(D!L1:L6,D!M1:M6),1,1)', 'INDEX(LINEST(D!B1:B8,D!D1:D8),1,1)',
    ],
    LOGEST: ['INDEX(LOGEST(1E+3,D!BC1:BC6),1,1)', 'INDEX(LOGEST(1E+3,D!BC1:BC6),1,7)', 'COLUMNS(LOGEST(1E+3,D!BC1:BC6))', 'INDEX(LOGEST(5,{1,2,3}),1,4)', 'LOGEST(D!L1:N6,D!BC1:BC6)'],
    // batch 2: an error among the new x is #VALUE!
    TREND: ['INDEX(TREND({1;2;3;5},{1,2;2,4;3,6;4,8}),4)', 'TREND(5,{1,2,3},{1,2,3})', 'INDEX(TREND({1;2;3;5},{1,2;2,4;3,6;4,8},{5,10}),1,1)', 'TREND(D!AY1:AY5,D!AZ1:AZ5,D!E1:E8)', 'TREND(D!AY1:AY5,D!AZ1:AZ5,{1,#N/A})'],
    GROWTH: ['GROWTH(5,{1,2,3},{1,2,3})', 'GROWTH(D!BB1:BB6,D!BC1:BC6,{17,#N/A})'],
  }
}

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [fn, fs] of Object.entries({ ...CASES, ...fits() })) add(fn, fn, ...fs)
  return out
}
