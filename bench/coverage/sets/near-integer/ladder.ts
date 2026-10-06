/** The first run: every context over the ladder of relative distances, below and above, positive and negative. */
import type { Add } from '../arrays/common.js'
import type { values, Value } from './values.js'

type Groups = ReturnType<typeof values>['g']

/** Twelve text characters, numbers 1…12 and a table of numbers for the whole-number arguments. */
export const S = '"abcdefghijkl"'
export const TWELVE = '{1,2,3,4,5,6,7,8,9,10,11,12}'
const CHOICES = Array.from({ length: 12 }, (_, i) => i + 1).join(',')

/** Cases of one group: a formula per value. */
export const eachOf = (add: Add) => (group: string, fn: string, vs: Value[], f: (x: string, v: Value) => string) => add(group, fn, ...vs.map((v) => f(v.at, v)))

export function ladder(add: Add, g: Groups) {
  const each = eachOf(add)
  const below = [...g.small, ...g.big, ...g.date]
  // the values as Excel holds them: the distance back in units of N·2^-j is 1
  each('exact', 'exact', g.small.filter((v) => v.n <= 2), (x, v) => `(${v.n}-${x})/${v.n}*2^${v.j}`)
  // INT and its kin, MOD, comparisons, criteria and lookups at every size
  each('int', 'INT', below, (x) => `INT(${x})`)
  each('trunc', 'TRUNC', below, (x) => `TRUNC(${x})`)
  each('rounddown', 'ROUNDDOWN', below, (x) => `ROUNDDOWN(${x},0)`)
  each('mod', 'MOD', below, (x) => `MOD(${x},1)`)
  each('eq', '=', below, (x, v) => `${x}=${v.n}`)
  each('lt', '<', below, (x, v) => `${x}<${v.n}`)
  each('countif', 'COUNTIF', below, (x, v) => `COUNTIF(${x},${v.n})`)
  each('match', 'MATCH', below, (x, v) => `MATCH(${v.n},${x},0)`)
  each('text', '&', g.small.filter((v) => v.n <= 2), (x) => `"a"&${x}`)
  each('day', 'DAY', g.date, (x) => `DAY(${x})`)
  each('weekday', 'WEEKDAY', g.date, (x) => `WEEKDAY(${x})`)
  // whole-number arguments of small size
  const small = g.small
  each('left', 'LEFT', small, (x) => `LEFT(${S},${x})`)
  each('right', 'RIGHT', small, (x) => `RIGHT(${S},${x})`)
  each('mid:start', 'MID', small, (x) => `MID(${S},${x},1)`)
  each('mid:len', 'MID', small, (x) => `MID(${S},1,${x})`)
  each('rept', 'REPT', small, (x) => `LEN(REPT("a",${x}))`)
  each('index', 'INDEX', small, (x) => `INDEX(${TWELVE},1,${x})`)
  each('choose', 'CHOOSE', small, (x) => `CHOOSE(${x},${CHOICES})`)
  each('round', 'ROUND', small, (x) => `ROUND(1.23456789012345,${x})`)
  each('large', 'LARGE', small, (x) => `LARGE(${TWELVE},${x})`)
  each('fact', 'FACT', small, (x) => `FACT(${x})`)
  each('dec2bin', 'DEC2BIN', small, (x) => `DEC2BIN(${x})`)
  each('address', 'ADDRESS', small, (x) => `ADDRESS(${x},1)`)
  each('bitand', 'BITAND', small, (x) => `BITAND(${x},15)`)
  each('iseven', 'ISEVEN', small, (x) => `ISEVEN(${x})`)
  each('combin', 'COMBIN', small, (x) => `COMBIN(12,${x})`)
  each('edate', 'EDATE', small, (x) => `EDATE(45688,${x})`)
  each('date', 'DATE', small, (x) => `DATE(2025,1,${x})`)
  each('vlookup', 'VLOOKUP', small.filter((v) => v.n <= 3), (x) => `VLOOKUP(13,D!L1:N6,${x},FALSE)`)
  each('dsum', 'DSUM', small.filter((v) => v.n <= 3), (x) => `DSUM(D!U1:Y7,${x},D!AG1:AG2)`)
  // above an integer: rounding up
  each('roundup', 'ROUNDUP', g.above, (x) => `ROUNDUP(${x},0)`)
  each('ceiling', 'CEILING', g.above, (x) => `CEILING(${x},1)`)
  each('even', 'EVEN', g.above, (x) => `EVEN(${x})`)
  each('eq:above', '=', g.above, (x, v) => `${x}=${v.n}`)
  each('gt:above', '>', g.above, (x, v) => `${x}>${v.n}`)
  // negative
  each('int:neg', 'INT', g.away, (x) => `INT(${x})`)
  each('eq:neg', '=', g.away, (x, v) => `${x}=${v.n}`)
  each('trunc:neg', 'TRUNC', g.toward, (x) => `TRUNC(${x})`)
  each('rounddown:neg', 'ROUNDDOWN', g.toward, (x) => `ROUNDDOWN(${x},0)`)
  each('eq:neg', '=', g.toward, (x, v) => `${x}=${v.n}`)
  // not an integer: comparisons near 0.3
  each('eq:0.3', '=', g.point3, (x) => `${x}=0.3`)
  each('lt:0.3', '<', g.point3, (x) => `${x}<0.3`)
  each('countif:0.3', 'COUNTIF', g.point3, (x) => `COUNTIF(${x},0.3)`)
  each('match:0.3', 'MATCH', g.point3, (x) => `MATCH(0.3,${x},0)`)
  // literals (Excel keeps 15 significant digits of them): the random check's cases
  add('literal', 'literal', 'INDEX({1,2,3},1,1.99999999999999)', 'LEFT("abc",0.999999999999999)', 'MID("abcdef",1.99999999999999,2)', '0.1+0.2=0.3', 'INT(FTEST(D!C1:C8,D!C1:C8))', 'FTEST(D!C1:C8,D!C1:C8)=1', '(1-FTEST(D!C1:C8,D!C1:C8))*2^52')
}
