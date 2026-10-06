/**
 * Formulas whose local text Excel is asked for (`snapshot.ts`): a call of every function of
 * `bench/coverage/excel-functions.json` that the coverage measurement has no case for, literals (TRUE and
 * FALSE, error literals, numbers, array constants, sheet names, strings with separators), then the
 * measurement's cases and the probes of `docs/excel-probes.json`. Formulas are in Excel's file form
 * (`_xlfn.` prefixes, `_xlpm.` LAMBDA parameters). Calls that reach outside the book (cube functions,
 * web services, add-ins) sit in a branch IF never computes, so they cannot block Excel with a prompt.
 */
import { readFileSync } from 'node:fs'
import type { Probe, SheetCells } from '../coverage/probes.js'
import { PROBE_SHEETS } from '../coverage/probes.js'

const ROOT = new URL('../../', import.meta.url)
const json = <T>(path: string): T => JSON.parse(readFileSync(new URL(path, ROOT), 'utf8')) as T

/** Sheets of the book besides the data sheet `D`: the probes' sheets and two with names that need quoting or are Cyrillic. */
export const SHEETS: SheetCells = { ...PROBE_SHEETS, 'Мой лист': { A1: 1, B2: 2 }, Данные: { A1: 3 } }

const never = (f: string) => `IF(FALSE,${f},0)`
const lambda = (params: string[], body: string) => `_xlfn.LAMBDA(${params.map((p) => `_xlpm.${p}`).join(',')},${body.replace(/\b([a-z])\b/g, '_xlpm.$1')})`

/**
 * Functions Excel deletes from a workbook it loads (they call code outside Excel), and in some books asks
 * to repair the book for them: their local names cannot be read from a book.
 */
export const NOT_IN_BOOKS = ['CALL', 'REGISTER.ID']

/** A call of each function without a case in the measurement. */
const CALLS: Record<string, string> = {
  ACCRINTM: 'ACCRINTM(39539,39614,0.1,1000)',
  ADDRESS: 'ADDRESS(2,3)',
  AMORDEGRC: 'AMORDEGRC(2400,39679,39813,300,1,0.15)',
  AMORLINC: 'AMORLINC(2400,39679,39813,300,1,0.15)',
  AREAS: 'AREAS(D!A1:B2)',
  ARRAYTOTEXT: '_xlfn.ARRAYTOTEXT(D!B1:B8)',
  ASC: 'ASC("a")',
  BAHTTEXT: 'BAHTTEXT(1)',
  BYCOL: `_xlfn.BYCOL(D!B1:C2,${lambda(['c'], 'SUM(c)')})`,
  BYROW: `_xlfn.BYROW(D!B1:C2,${lambda(['r'], 'SUM(r)')})`,
  CELL: 'CELL("row",D!A1)',
  CONFIDENCE: 'CONFIDENCE(0.05,2.5,50)',
  COUPDAYBS: 'COUPDAYBS(40568,40862,2,1)',
  COUPDAYSNC: 'COUPDAYSNC(40568,40862,2,1)',
  COUPNCD: 'COUPNCD(40568,40862,2,1)',
  COUPNUM: 'COUPNUM(40568,40862,2,1)',
  COUPPCD: 'COUPPCD(40568,40862,2,1)',
  CUBEKPIMEMBER: never('CUBEKPIMEMBER("c","k",1)'),
  CUBEMEMBER: never('CUBEMEMBER("c","m")'),
  CUBEMEMBERPROPERTY: never('CUBEMEMBERPROPERTY("c","m","p")'),
  CUBERANKEDMEMBER: never('CUBERANKEDMEMBER("c","s",1)'),
  CUBESET: never('CUBESET("c","s")'),
  CUBESETCOUNT: never('CUBESETCOUNT("s")'),
  CUBEVALUE: never('CUBEVALUE("c")'),
  DBCS: '_xlfn.DBCS("a")',
  DETECTLANGUAGE: never('_xlfn.DETECTLANGUAGE("a")'),
  DURATION: 'DURATION(39448,42370,0.08,0.09,2)',
  ENCODEURL: '_xlfn.ENCODEURL("a b")',
  'ERF.PRECISE': '_xlfn.ERF.PRECISE(1)',
  'ERFC.PRECISE': '_xlfn.ERFC.PRECISE(1)',
  EUROCONVERT: never('EUROCONVERT(1,"EUR","DEM")'),
  FILTER: '_xlfn._xlws.FILTER(D!B1:B8,D!B1:B8>0)',
  FILTERXML: never('_xlfn.FILTERXML("<a>1</a>","//a")'),
  FINDB: 'FINDB("b","abc")',
  'FORECAST.ETS': never('_xlfn.FORECAST.ETS(9,D!B1:B8,D!B1:B8)'),
  'FORECAST.ETS.CONFINT': never('_xlfn.FORECAST.ETS.CONFINT(9,D!B1:B8,D!B1:B8)'),
  'FORECAST.ETS.SEASONALITY': never('_xlfn.FORECAST.ETS.SEASONALITY(D!B1:B8,D!B1:B8)'),
  'FORECAST.ETS.STAT': never('_xlfn.FORECAST.ETS.STAT(D!B1:B8,D!B1:B8,1)'),
  'FORECAST.LINEAR': '_xlfn.FORECAST.LINEAR(1,D!B1:B8,D!B1:B8)',
  FORMULATEXT: '_xlfn.FORMULATEXT(D!A1)',
  GETPIVOTDATA: never('GETPIVOTDATA("a",D!A1)'),
  GROUPBY: never(`_xlfn.GROUPBY(D!F1:F7,D!B1:B7,${lambda(['x'], 'SUM(x)')})`),
  HYPERLINK: 'HYPERLINK("https://example.com","a")',
  IMAGE: never('_xlfn.IMAGE("https://example.com/a.png")'),
  INDIRECT: 'INDIRECT("D!A1")',
  INFO: 'INFO("osversion")',
  INTRATE: 'INTRATE(39493,39583,1000000,1014420)',
  ISFORMULA: '_xlfn.ISFORMULA(D!A1)',
  ISOMITTED: `${lambda(['x'], '_xlfn.ISOMITTED(x)')}(1)`,
  ISREF: 'ISREF(D!A1)',
  JIS: 'JIS("a")',
  LAMBDA: `${lambda(['x'], 'x+1')}(1)`,
  LEFTB: 'LEFTB("abc",1)',
  LENB: 'LENB("abc")',
  LET: '_xlfn.LET(_xlpm.x,1,_xlpm.x+1)',
  MAKEARRAY: `_xlfn.MAKEARRAY(2,2,${lambda(['r', 'c'], 'r*c')})`,
  MAP: `_xlfn.MAP(D!B1:B8,${lambda(['x'], 'x*2')})`,
  MDETERM: 'MDETERM({1,2;3,4})',
  MDURATION: 'MDURATION(39448,42370,0.08,0.09,2)',
  MIDB: 'MIDB("abc",1,1)',
  MINVERSE: 'MINVERSE({1,2;3,4})',
  MODE: 'MODE(1,2,2)',
  NOW: 'NOW()',
  ODDFPRICE: 'ODDFPRICE(39763,44256,39736,39873,0.0785,0.0625,100,2)',
  ODDFYIELD: 'ODDFYIELD(39763,44256,39736,39873,0.0575,84.5,100,2)',
  ODDLPRICE: 'ODDLPRICE(39494,39600,39388,0.0375,0.0405,100,2)',
  ODDLYIELD: 'ODDLYIELD(39494,39600,39388,0.0375,99.875,100,2)',
  OFFSET: 'OFFSET(D!A1,1,1)',
  PERCENTILE: 'PERCENTILE(D!B1:B8,0.5)',
  PERCENTOF: '_xlfn.PERCENTOF(1,2)',
  PERCENTRANK: 'PERCENTRANK(D!B1:B8,8)',
  PHONETIC: 'PHONETIC(D!A3)',
  PIVOTBY: never(`_xlfn.PIVOTBY(D!F1:F7,D!F1:F7,D!B1:B7,${lambda(['x'], 'SUM(x)')})`),
  POISSON: 'POISSON(2,5,TRUE)',
  PRICE: 'PRICE(39493,43054,0.0575,0.065,100,2)',
  PRICEMAT: 'PRICEMAT(39493,39583,39431,0.061,0.061)',
  QUARTILE: 'QUARTILE(D!B1:B8,1)',
  RAND: 'RAND()',
  RANDARRAY: '_xlfn.RANDARRAY(2)',
  RANDBETWEEN: 'RANDBETWEEN(1,2)',
  RANK: 'RANK(8,D!B1:B8)',
  RECEIVED: 'RECEIVED(39493,39583,1000000,0.0575)',
  REDUCE: `_xlfn.REDUCE(0,D!B1:B8,${lambda(['a', 'b'], 'a+b')})`,
  REGEXEXTRACT: '_xlfn.REGEXEXTRACT("abc","b")',
  REGEXREPLACE: '_xlfn.REGEXREPLACE("abc","b","x")',
  REGEXTEST: '_xlfn.REGEXTEST("abc","b")',
  REPLACEB: 'REPLACEB("abc",1,1,"x")',
  RIGHTB: 'RIGHTB("abc",1)',
  RTD: never('RTD("a","","b")'),
  SCAN: `_xlfn.SCAN(0,D!B1:B8,${lambda(['a', 'b'], 'a+b')})`,
  SEARCHB: 'SEARCHB("b","abc")',
  SEQUENCE: '_xlfn.SEQUENCE(2)',
  SHEET: '_xlfn.SHEET()',
  SHEETS: '_xlfn.SHEETS()',
  SORTBY: '_xlfn.SORTBY(D!B1:B8,D!B1:B8)',
  STDEV: 'STDEV(1,2)',
  STOCKHISTORY: never('_xlfn.STOCKHISTORY("MSFT",45000)'),
  TEXTAFTER: '_xlfn.TEXTAFTER("a-b","-")',
  TEXTBEFORE: '_xlfn.TEXTBEFORE("a-b","-")',
  TEXTSPLIT: '_xlfn.TEXTSPLIT("a b"," ")',
  TOCOL: '_xlfn.TOCOL(D!B1:C2)',
  TODAY: 'TODAY()',
  TOROW: '_xlfn.TOROW(D!B1:C2)',
  TRANSLATE: never('_xlfn.TRANSLATE("a")'),
  TRIMRANGE: '_xlfn.TRIMRANGE(D!B1:B8)',
  VALUETOTEXT: '_xlfn.VALUETOTEXT(1)',
  VAR: 'VAR(1,2)',
  VDB: 'VDB(2400,300,10,0,1)',
  WEBSERVICE: never('_xlfn.WEBSERVICE("https://example.com")'),
  WEIBULL: 'WEIBULL(105,20,100,TRUE)',
  WRAPCOLS: '_xlfn.WRAPCOLS(D!B1:B8,2)',
  WRAPROWS: '_xlfn.WRAPROWS(D!B1:B8,2)',
  XLOOKUP: '_xlfn.XLOOKUP(8,D!B1:B8,D!B1:B8)',
  XMATCH: '_xlfn.XMATCH(8,D!B1:B8)',
  YIELD: 'YIELD(39494,42998,0.0575,95.04287,100,2)',
  YIELDDISC: 'YIELDDISC(39494,39651,99.795,100)',
  YIELDMAT: 'YIELDMAT(39494,39651,39377,0.0625,100.0123)',
}

/** Error literals Excel 16 has had for years; the newer ones go in a part of their own (`NEW_ERRORS`). */
export const ERRORS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A']
export const NEW_ERRORS = ['#GETTING_DATA', '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!', '#PYTHON!']

/** Literals and syntax: what a locale may write differently, apart from function names. */
const LITERALS = [
  'TRUE',
  'FALSE',
  'AND(TRUE,FALSE,true)',
  '1.5',
  '-0.25+.5',
  '1.5E-07*2E+21',
  '50%+0.5%',
  'SUM({1,2;3,4})',
  'SUM({1.5,-2.5;3,4})',
  'INDEX({"a","b";"c,d","e;f"},2,1)',
  'AND({TRUE,FALSE;FALSE,TRUE})',
  'ISNA({#N/A,1})',
  'ROWS({1;2;3})',
  'SUM((D!B1,D!B2),D!B3)',
  'SUM(D!B1:B3 D!B2:B4)',
  '"a,b;c.d\\e"&"1,5"',
  "'Мой лист'!A1+Данные!A1+'Мой лист'!B2",
  "SUM('Мой лист'!A1:B2,Данные!A:A,Данные!1:1)",
  'D!$A$1+D!A$1+D!$A1',
  '2^3&"x"',
  'IF(1<>2,1>=2,1<=2)',
  '-D!B1',
]

/** Function names of the measurement's cases, with their prefixes stripped. */
const called = (f: string) => [...f.matchAll(/(?:_xlfn\.)?(?:_xlws\.)?([A-Z][A-Z0-9.]*)\(/g)].map((m) => m[1])

export interface Probes {
  /** Parts sent to Excel one by one (the newer error literals alone: Excel may reject the book). */
  parts: Probe[][]
  /** Functions of the reference list. */
  functions: string[]
}

export function buildProbes(): Probes {
  const functions = json<{ name: string }[]>('bench/coverage/excel-functions.json').map((f) => f.name)
  const measured = json<{ cases: { id: string; fn: string; formula: string }[] }>('docs/coverage-formulajs/cases.json').cases
  const probes = json<{ cases: Probe[] }>('docs/excel-probes.json').cases
  const inMeasurement = new Set(measured.flatMap((c) => called(c.formula)))
  const missing = functions.filter((f) => !inMeasurement.has(f) && !NOT_IN_BOOKS.includes(f))
  const unknown = missing.filter((f) => !(f in CALLS))
  if (unknown.length) throw new Error(`no call for ${unknown.join(', ')}`)
  const own: Probe[] = [
    ...missing.map((f) => ({ id: `fn:${f}`, group: 'function', fn: f, formula: CALLS[f] })),
    ...LITERALS.map((f, i) => ({ id: `lit:${i + 1}`, group: 'literal', formula: f })),
    ...ERRORS.map((e) => ({ id: `err:${e}`, group: 'error', formula: e })),
  ]
  const rest = [...measured.map((c) => ({ id: `m:${c.id}`, group: 'measurement', fn: c.fn, formula: c.formula })), ...probes.map((p) => ({ ...p, id: `p:${p.id}` }))]
  const parts = [own, NEW_ERRORS.map((e) => ({ id: `err:${e}`, group: 'error', formula: e }))]
  for (let i = 0; i < rest.length; i += 400) parts.push(rest.slice(i, i + 400))
  return { parts, functions }
}
