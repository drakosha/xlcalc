/**
 * The third run: the size of the tolerance of whole-number arguments (between 2·10^-7 and 3·10^-7 below an
 * integer) and of ISEVEN and ISODD (about 10^-10), and the whole-number arguments the second run left out.
 */
import type { Add } from '../arrays/common.js'
import { eachOf, S, TWELVE } from './ladder.js'
import type { values } from './values.js'

type Groups = ReturnType<typeof values>['g']

/** The rest of the whole-number arguments, `@` in each formula (2^-40 and 2^-30 below the integers given). */
const SAMPLED: [string, string, number[]][] = [
  ['DATE', 'DATE(@,1,1)', [2025]], ['INDEX', 'INDEX(D!B1:B8,1,1,@)', [1]], ['TOROW', 'COLUMNS(TOROW(D!D4:D6,@))', [1]],
  ['PERCENTRANK.EXC', 'PERCENTRANK.EXC(D!B1:B8,10,@)', [2, 3]], ['PERCENTRANK', 'PERCENTRANK(D!B1:B8,10,@)', [2, 3]],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(@,0.5,1)', [2, 3]], ['CHISQ.DIST.RT', 'CHISQ.DIST.RT(1,@)', [2, 3]], ['F.DIST.RT', 'F.DIST.RT(1,@,3)', [2, 3]],
  ['T.DIST.2T', 'T.DIST.2T(1,@)', [2, 3]], ['T.DIST.RT', 'T.DIST.RT(1,@)', [2, 3]], ['CHISQ.INV', 'CHISQ.INV(0.5,@)', [2, 3]],
  ['CHISQ.INV.RT', 'CHISQ.INV.RT(0.5,@)', [2, 3]], ['CONFIDENCE.NORM', 'CONFIDENCE.NORM(0.05,1,@)', [2, 3]], ['F.INV', 'F.INV(0.5,@,3)', [2, 3]],
  ['F.INV.RT', 'F.INV.RT(0.5,@,3)', [2, 3]], ['T.INV', 'T.INV(0.75,@)', [2, 3]], ['T.INV.2T', 'T.INV.2T(0.5,@)', [2, 3]],
  ['T.TEST', 'T.TEST(D!B1:B8,D!C1:C8,2,@)', [2, 3]], ['TTEST', 'TTEST(D!B1:B8,D!C1:C8,@,1)', [2]], ['TTEST', 'TTEST(D!B1:B8,D!C1:C8,2,@)', [2, 3]],
  ['BINOMDIST', 'BINOMDIST(@,5,0.5,FALSE)', [2, 3]], ['CHIDIST', 'CHIDIST(1,@)', [2, 3]], ['CHIINV', 'CHIINV(0.5,@)', [2, 3]],
  ['CONFIDENCE', 'CONFIDENCE(0.05,1,@)', [2, 3]], ['CRITBINOM', 'CRITBINOM(@,0.5,0.9)', [2, 3]], ['BINOM.INV', 'BINOM.INV(@,0.5,0.9)', [2, 3]],
  ['FDIST', 'FDIST(1,@,3)', [2, 3]], ['FINV', 'FINV(0.5,@,3)', [2, 3]], ['HYPGEOMDIST', 'HYPGEOMDIST(@,4,5,10)', [2, 3]],
  ['NEGBINOMDIST', 'NEGBINOMDIST(@,2,0.5)', [2, 3]], ['POISSON', 'POISSON(@,2,FALSE)', [2, 3]], ['TDIST', 'TDIST(1,@,1)', [2, 3]],
  ['TDIST', 'TDIST(1,3,@)', [2]], ['TINV', 'TINV(0.5,@)', [2, 3]], ['PERMUTATIONA', 'PERMUTATIONA(5,@)', [2, 3]],
  ['FINDB', 'FINDB("a","aaaa",@)', [2, 3]], ['SEARCHB', 'SEARCHB("a","aaaa",@)', [2, 3]], ['REPLACEB', 'REPLACEB("abcd",@,1,"-")', [2, 3]],
  ['REPLACEB', 'REPLACEB("abcd",1,@,"-")', [2, 3]], ['RIGHTB', 'RIGHTB("abcd",@)', [2, 3]], ['MIDB', 'MIDB("abcd",1,@)', [2, 3]],
  ['TEXTSPLIT', 'COLUMNS(TEXTSPLIT("aXbxc","x",,,@))', [1]], ['TEXTBEFORE', 'TEXTBEFORE("aXb","x",1,@)', [1]], ['TEXTBEFORE', 'TEXTBEFORE("a-b","+",1,0,@)', [1]],
  ['TEXTAFTER', 'TEXTAFTER("aXb","x",1,@)', [1]], ['HYPGEOM.DIST', 'HYPGEOM.DIST(1,@,5,10,FALSE)', [2, 3]], ['NEGBINOM.DIST', 'NEGBINOM.DIST(1,@,0.5,FALSE)', [2, 3]],
  ['PERMUT', 'PERMUT(@,1)', [2, 3]], ['DB', 'DB(1000,100,@,1)', [2, 3]], ['SUBTOTAL', 'SUBTOTAL(@,D!B1:B8)', [5]], ['AGGREGATE', 'AGGREGATE(9,@,D!B1:B8)', [5]],
]

export function tolerance(add: Add, g: Groups) {
  const each = eachOf(add)
  const of = (ns: number[]) => g.tau.filter((v) => ns.includes(v.n))
  each('tau:left', 'LEFT', of([1, 2, 3, 10]), (x) => `LEFT(${S},${x})`)
  each('tau:index', 'INDEX', of([1, 2, 3, 10]), (x) => `INDEX(${TWELVE},1,${x})`)
  each('tau:round', 'ROUND', of([1, 2, 3, 10]), (x) => `ROUND(1.23456789012345,${x})`)
  each('tau:address', 'ADDRESS', of([1, 2, 3, 10]), (x) => `ADDRESS(${x},1)`)
  each('tau:date', 'DATE', of([1, 2, 3, 10]), (x) => `DATE(2025,1,${x})`)
  each('tau:fact', 'FACT', of([2, 3, 10]), (x) => `FACT(${x})`)
  each('tau:vlookup', 'VLOOKUP', of([1, 2, 3]), (x) => `VLOOKUP(13,D!L1:N6,${x},FALSE)`)
  each('tau:dsum', 'DSUM', of([1, 2, 3]), (x) => `DSUM(D!U1:Y7,${x},D!AG1:AG2)`)
  each('tau:subtotal', 'SUBTOTAL', of([9]), (x) => `SUBTOTAL(${x},D!B1:B8)`)
  each('tau:round', 'ROUND', g.tauNeg, (x) => `ROUND(123.4,${x})`)
  each('tau:date', 'DATE', g.tauNeg, (x) => `DATE(2025,1,${x})`)
  each('parity:ISEVEN', 'ISEVEN', g.parityFine, (x) => `ISEVEN(${x})`)
  each('parity:ISODD', 'ISODD', g.parityFine, (x) => `ISODD(${x})`)
  const sample = [...g.sample, ...g.sample2]
  for (const [fn, f, ns] of SAMPLED) each(`sample:${fn}`, fn, sample.filter((v) => ns.includes(v.n)), (x) => f.replace('@', x))
}
