/**
 * The second run, around the boundaries the ladder found: whole-number arguments near a tolerance of 10^-7,
 * which functions take a whole number with a tolerance and which truncate it, comparisons and lookups of every
 * kind, ISEVEN and ISODD, the date functions reading a moment to the second.
 */
import type { Add } from '../arrays/common.js'
import { eachOf, S, TWELVE } from './ladder.js'
import type { values, Value } from './values.js'

type Groups = ReturnType<typeof values>['g']

/** A whole-number argument per function: the formula and the integers whose values (2^-40 and 2^-30 below) it gets. */
const SAMPLED: [string, string, number[]][] = [
  ['WEEKDAY', 'WEEKDAY(45688,x)', [2, 3]], ['ROUNDDOWN', 'ROUNDDOWN(1.23456,x)', [2, 3]], ['ROUNDUP', 'ROUNDUP(1.23456,x)', [2, 3]],
  ['TRUNC', 'TRUNC(1.23456,x)', [2, 3]], ['FIXED', 'FIXED(1.23456,x)', [2, 3]], ['DOLLAR', 'DOLLAR(1.23456,x)', [2, 3]],
  ['FIND', 'FIND("a","aaaa",x)', [2, 3]], ['SEARCH', 'SEARCH("a","aaaa",x)', [2, 3]], ['REPLACE', 'REPLACE("abcd",x,1,"-")', [2, 3]],
  ['REPLACE', 'REPLACE("abcd",1,x,"-")', [2, 3]], ['SUBSTITUTE', 'SUBSTITUTE("aaaa","a","b",x)', [2, 3]], ['LEFTB', 'LEFTB("abcd",x)', [2, 3]],
  ['MIDB', 'MIDB("abcd",x,1)', [2, 3]], ['UNICHAR', 'UNICHAR(x+64)', [2, 3]], ['TEXTBEFORE', 'TEXTBEFORE("a-b-c-d","-",x)', [2, 3]],
  ['TEXTAFTER', 'TEXTAFTER("a-b-c-d","-",x)', [2, 3]], ['INDEX', 'INDEX({1,2,3;4,5,6;7,8,9},x,1)', [2, 3]], ['INDEX', 'INDEX(D!B1:B8,x)', [2, 3]],
  ['ADDRESS', 'ADDRESS(1,x)', [2, 3]], ['ADDRESS', 'ADDRESS(1,1,x)', [2, 3]], ['DATE', 'DATE(2025,x,1)', [2, 3]],
  ['PERMUT', 'PERMUT(5,x)', [2, 3]], ['BINOM.DIST', 'BINOM.DIST(x,5,0.5,FALSE)', [2, 3]], ['BINOM.DIST', 'BINOM.DIST(1,x,0.5,FALSE)', [2, 3]],
  ['POISSON.DIST', 'POISSON.DIST(x,2,FALSE)', [2, 3]], ['HYPGEOM.DIST', 'HYPGEOM.DIST(x,4,5,10,FALSE)', [2, 3]], ['NEGBINOM.DIST', 'NEGBINOM.DIST(x,2,0.5,FALSE)', [2, 3]],
  ['CHISQ.DIST', 'CHISQ.DIST(1,x,TRUE)', [2, 3]], ['T.DIST', 'T.DIST(1,x,TRUE)', [2, 3]], ['F.DIST', 'F.DIST(1,x,3,TRUE)', [2, 3]],
  ['T.TEST', 'T.TEST(D!B1:B8,D!C1:C8,x,1)', [2]], ['CONFIDENCE.T', 'CONFIDENCE.T(0.05,1,x)', [2, 3]], ['BINOM.INV', 'BINOM.INV(x,0.5,0.5)', [2, 3]],
  ['PERCENTRANK.INC', 'PERCENTRANK.INC(D!B1:B8,10,x)', [2, 3]], ['CEILING.MATH', 'CEILING.MATH(-2.5,1,x)', [1]], ['FLOOR.MATH', 'FLOOR.MATH(-2.5,1,x)', [1]],
  ['DB', 'DB(1000,100,5,1,x)', [2, 3]], ['TOCOL', 'ROWS(TOCOL(D!D4:D6,x))', [1]], ['ARRAYTOTEXT', 'ARRAYTOTEXT({1,"a"},x)', [1]],
  ['VALUETOTEXT', 'VALUETOTEXT("a",x)', [1]], ['MATCH', 'MATCH(4,D!L1:L6,x)', [1]], ['HLOOKUP', 'HLOOKUP(5,D!L8:Q10,x,FALSE)', [2, 3]],
  ['SUBTOTAL', 'SUBTOTAL(x,D!B1:B8)', [9]], ['AGGREGATE', 'AGGREGATE(x,0,D!B1:B8)', [9]], ['SMALL', 'SMALL(D!B1:B8,x)', [2, 3]],
  ['QUARTILE.INC', 'QUARTILE.INC(D!B1:B8,x)', [2, 3]], ['RANK.EQ', 'RANK.EQ(4,D!B1:B8,x)', [1]], ['ROMAN', 'ROMAN(x)', [2, 3]],
  ['BASE', 'BASE(5,x)', [2, 3]], ['BASE', 'BASE(5,2,x)', [2, 3]], ['DEC2BIN', 'DEC2BIN(5,x)', [3, 4]], ['BITLSHIFT', 'BITLSHIFT(1,x)', [2, 3]],
  ['EOMONTH', 'EOMONTH(45688,x)', [2, 3]], ['WORKDAY', 'WORKDAY(45688,x)', [2, 3]], ['NETWORKDAYS.INTL', 'NETWORKDAYS.INTL(45688,45700,x)', [2, 3]],
  ['COUPNUM', 'COUPNUM(45688,46000,x)', [2]], ['YEARFRAC', 'YEARFRAC(45688,46000,x)', [2, 3]], ['GCD', 'GCD(x,4)', [2, 3]],
  ['FACTDOUBLE', 'FACTDOUBLE(x)', [2, 3]], ['CHOOSECOLS', 'CHOOSECOLS({1,2,3},x)', [2, 3]], ['ISODD', 'ISODD(x)', [2, 3]],
  ['FACT', 'FACT(x)', [2, 3]], ['LEFT', `LEFT(${S},x)`, [2, 3]], ['MID', `MID(${S},x,1)`, [2, 3]], ['CHOOSE', 'CHOOSE(x,1,2,3,4)', [2, 3]],
]

/** Comparisons and lookups of a value near 2 against 2. */
const COMPARED: [string, string][] = [
  ['<=', 'x<=2'], ['>=', 'x>=2'], ['<>', 'x<>2'], ['>', 'x>2'], ['SUMIF', 'SUMIF(x,2)'], ['COUNTIFS', 'COUNTIFS(x,2)'],
  ['COUNTIF', 'COUNTIF(x,"<"&2)'], ['COUNTIF', 'COUNTIF(x,">="&2)'], ['MAXIFS', 'MAXIFS(x,x,2)'], ['MATCH', 'MATCH(x,{2},1)'],
  ['VLOOKUP', 'VLOOKUP(x,{2,"b"},2,TRUE)'], ['VLOOKUP', 'VLOOKUP(x,{2,"b"},2,FALSE)'], ['XMATCH', 'XMATCH(x,{2})'], ['XLOOKUP', 'XLOOKUP(x,{2},{"b"})'],
  ['UNIQUE', 'ROWS(UNIQUE(VSTACK(x,2)))'], ['RANK.EQ', 'RANK.EQ(2,x)'], ['SUMPRODUCT', 'SUMPRODUCT(--(x={2}))'], ['DELTA', 'DELTA(x,2)'],
  ['GESTEP', 'GESTEP(x,2)'], ['EXACT', 'EXACT(x,2)'],
]

/** The date functions over a moment just before midnight. */
const DATED: [string, string][] = [
  ['DAY', 'DAY(x)'], ['WEEKDAY', 'WEEKDAY(x)'], ['MONTH', 'MONTH(x)'], ['YEAR', 'YEAR(x)'], ['EOMONTH', 'EOMONTH(x,0)'], ['DAYS', 'DAYS(45700,x)'],
  ['DATEDIF', 'DATEDIF(x,45700,"d")'], ['WORKDAY', 'WORKDAY(x,1)'], ['WEEKNUM', 'WEEKNUM(x)'], ['HOUR', 'HOUR(x)'], ['SECOND', 'SECOND(x)'],
  ['EDATE', 'EDATE(x,1)'], ['NETWORKDAYS', 'NETWORKDAYS(x,45700)'], ['TEXT', 'TEXT(x,"d")'],
]

export function fine(add: Add, g: Groups) {
  const each = eachOf(add)
  const of = (ns: number[]) => g.fine.filter((v) => ns.includes(v.n))
  // the tolerance of whole-number arguments, around 10^-7 relative
  each('fine:left', 'LEFT', of([1, 2, 3, 10]), (x) => `LEFT(${S},${x})`)
  each('fine:right', 'RIGHT', of([1, 2, 3, 10]), (x) => `RIGHT(${S},${x})`)
  each('fine:index', 'INDEX', of([1, 2, 3, 10]), (x) => `INDEX(${TWELVE},1,${x})`)
  each('fine:round', 'ROUND', of([1, 2, 3, 10]), (x) => `ROUND(1.23456789012345,${x})`)
  each('fine:fact', 'FACT', of([2, 3, 10]), (x) => `FACT(${x})`)
  each('fine:address', 'ADDRESS', g.fine, (x) => `ADDRESS(${x},1)`)
  each('fine:date', 'DATE', of([1, 2, 3, 10, 100]), (x) => `DATE(2025,1,${x})`)
  each('fine:vlookup', 'VLOOKUP', of([1, 2, 3]), (x) => `VLOOKUP(13,D!L1:N6,${x},FALSE)`)
  each('fine:dsum', 'DSUM', of([1, 2, 3]), (x) => `DSUM(D!U1:Y7,${x},D!AG1:AG2)`)
  // which functions take a whole number with that tolerance: 2^-40 and 2^-30 below
  for (const [fn, f, ns] of SAMPLED) each(`sample:${fn}`, fn, g.sample.filter((v) => ns.includes(v.n)), (x) => f.replace('x', x))
  // comparisons and lookups: below and above 2, within and past 15 digits
  const near2: Value[] = [...g.small.filter((v) => v.n === 2 && [53, 51, 49, 48].includes(v.j!)), ...g.above.filter((v) => v.n === 2 && [52, 49, 48].includes(v.j!))]
  for (const [fn, f] of COMPARED) each(`cmp:${fn}`, fn, near2, (x) => f.replaceAll('x', x))
  // ISEVEN and ISODD
  each('parity:ISEVEN', 'ISEVEN', g.parity, (x) => `ISEVEN(${x})`)
  each('parity:ISODD', 'ISODD', g.parity, (x) => `ISODD(${x})`)
  // a moment to the second
  for (const [fn, f] of DATED) each(`second:${fn}`, fn, g.seconds, (x) => f.replace('x', x))
}
