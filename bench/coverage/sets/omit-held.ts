/**
 * Held-out cases for omitted arguments: each answer was predicted from the rules the `omit` set found and
 * committed before Excel computed the case (`refused`: Excel does not load the formula). The rules: a left-out
 * argument is a blank, read by kind as 0, "" or FALSE; #VALUE! where the function takes an array or a reference;
 * the default where the function treats it as not given (the newer functions' optional arguments, PRODUCT's
 * values, TEXTJOIN's ignore_empty as TRUE); in the former Analysis ToolPak's functions not given, a required one
 * #N/A; not loaded where Excel takes only a reference. Answers: `test/excel/cases/omit-held.json`.
 */
import type { Val } from '../formulajs.js'
import type { Probe } from '../probes.js'
import { xl } from './arrays/common.js'
import { n, s, b, e } from '../vals.js'

/** Formula (the engine's spelling), the function it checks and the predicted answer. */
const CASES: [string, string, Val | 'refused'][] = [
  ['RANDBETWEEN(,0)', 'RANDBETWEEN', e('NA')],
  ['ISO.CEILING(,2)', 'ISO.CEILING', n(0)],
  ['CONVERT(,"m","km")', 'CONVERT', e('NA')],
  ['EOMONTH(45688,)', 'EOMONTH', e('NA')],
  ['DOLLARDE(1.02,)', 'DOLLARDE', e('NA')],
  ['GCD(,4)', 'GCD', e('NA')],
  ['LCM(6,,4)', 'LCM', n(12)],
  ['IMSUM("1+i",,"2")', 'IMSUM', s('3+i')],
  ['WORKDAY(DATE(2025,1,1),,D!AP1:AP8)', 'WORKDAY', e('NA')],
  ['YEARFRAC(45658,45688,)', 'YEARFRAC', n(30 / 360)],
  ['COUNT(,,)', 'COUNT', n(3)],
  ['COUNTA(,,,)', 'COUNTA', n(4)],
  ['AND(,TRUE)', 'AND', b(false)],
  ['CONCAT(,,"x")', 'CONCAT', s('x')],
  ['SUM(,,5)', 'SUM', n(5)],
  ['AVERAGE(,,6)', 'AVERAGE', n(2)],
  ['MAX(,,-5)', 'MAX', n(0)],
  ['STDEV(,,4)', 'STDEV', n(2.309401076758503)],
  ['NPV(0.1,,100)', 'NPV', n(100 / 1.21)],
  ['IF(,,)', 'IF', n(0)],
  ['IFERROR(,)', 'IFERROR', n(0)],
  ['IFS(FALSE,1,,2)', 'IFS', e('NA')],
  ['SWITCH(1,,"a")', 'SWITCH', e('NA')],
  ['CHOOSE(3,1,2,)', 'CHOOSE', n(0)],
  ['SUM(IF({1,0},,5))', 'IF', n(5)],
  ['COUNT(IF(TRUE,))', 'COUNT', n(1)],
  ['COUNTA(IF(TRUE,))', 'COUNTA', n(1)],
  ['PEARSON(,)', 'PEARSON', e('VALUE')],
  ['INDEX(,1)', 'INDEX', e('VALUE')],
  ['SUMPRODUCT(D!B1:B2,,)', 'SUMPRODUCT', e('VALUE')],
  ['HSTACK(1,)', 'HSTACK', e('VALUE')],
  ['TEXTJOIN(,,"a","","b")', 'TEXTJOIN', s('ab')],
  ['PRODUCT(,3)', 'PRODUCT', n(3)],
  ['PRODUCT(,)', 'PRODUCT', n(0)],
  ['TEXTBEFORE("abc","-",,,,)', 'TEXTBEFORE', e('NA')],
  ['ARRAYTOTEXT(TEXTSPLIT("a,b;c",",",";",,,),1)', 'TEXTSPLIT', s('{"a","b";"c",#N/A}')],
  ['ARRAYTOTEXT(SORTBY(D!B1:B4,D!C1:C4,1,D!B1:B4,),1)', 'SORTBY', s('{4;8;15;16}')],
  ['ARRAYTOTEXT(RANDARRAY(1,2,1,,TRUE),1)', 'RANDARRAY', s('{1,1}')],
  ['ARRAYTOTEXT(EXPAND({1,2},,,9),1)', 'EXPAND', s('{1,2}')],
  ['ARRAYTOTEXT(TAKE(D!B1:C4,,),1)', 'TAKE', s('{4,2;8,7;15,11;16,19}')],
  ['ARRAYTOTEXT(SEQUENCE(,,,),1)', 'SEQUENCE', s('{1}')],
  ['XLOOKUP(5,D!L1:L6,D!M1:M6,,,)', 'XLOOKUP', s('five')],
  ['MATCH(8,D!L1:L6,)', 'MATCH', n(5)],
  ['VLOOKUP(8,D!L1:N6,3,)', 'VLOOKUP', n(80)],
  ['ROUNDUP(1.21,)', 'ROUNDUP', n(2)],
  ['TRUNC(1.29,)', 'TRUNC', n(1)],
  ['MID("abcdef",,)', 'MID', e('VALUE')],
  ['DATE(,,)', 'DATE', e('NUM')],
  ['CEILING.MATH(-2.5,,)', 'CEILING.MATH', n(-2)],
  ['FIXED(1234.567,,TRUE)', 'FIXED', s('1234.57')],
  ['ADDRESS(2,3,,,"S")', 'ADDRESS', s('S!$C$2')],
  ['TEXT(,)', 'TEXT', s('')],
  ['REPT(,)', 'REPT', s('')],
  ['COUNTIF(,1)', 'COUNTIF', 'refused'],
  ['SUMIFS(,D!B1:B8,1)', 'SUMIFS', 'refused'],
  ['AVERAGEIF(D!B1:B8,">5",)', 'AVERAGEIF', 'refused'],
  ['COUNTIFS(D!B1:B8,">5",,1)', 'COUNTIFS', 'refused'],
  ['RANK(1,,0)', 'RANK', 'refused'],
  ['SUBTOTAL(9,)', 'SUBTOTAL', 'refused'],
  ['AGGREGATE(9,0,)', 'AGGREGATE', 'refused'],
  ['DSUM(D!U1:Y7,"Height",)', 'DSUM', 'refused'],
]

/** Of the cases, those Excel does not load: an argument left out where Excel takes only a reference. */
export const REFUSED: readonly string[] = ['COUNTIF(,1)', 'SUMIFS(,D!B1:B8,1)', 'AVERAGEIF(D!B1:B8,">5",)', 'COUNTIFS(D!B1:B8,">5",,1)', 'RANK(1,,0)', 'SUBTOTAL(9,)']

/** The cases Excel loads; a case keeps its number among all of them. */
export function buildProbes(): Probe[] {
  return CASES.flatMap(([f, fn], i) => (REFUSED.includes(f) ? [] : [{ id: `held#${i + 1}`, group: 'held', fn, formula: xl(f) }]))
}

/** The predicted answers by case id. */
export const PREDICTED: Readonly<Record<string, Val | 'refused'>> = Object.fromEntries(CASES.map(([, , v], i) => [`held#${i + 1}`, v]))
