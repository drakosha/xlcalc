/**
 * Two rules the engine wrote twice, held out. `BINOM.DIST.RANGE` past 170 trials, where its factorials overflowed to
 * #NUM! while `BINOM.DIST` computes: the rule predicted is the sum of the binomial densities from the first count to
 * the second, `BINOM.DIST`'s own, up to 2^31 − 2 trials. `SUMX2MY2`, `SUMXMY2` and `SUMX2PY2` given an error in one cell
 * against a range of another size, where the engine answered #N/A (the sizes first) and `CORREL` the cell's error: the
 * rule predicted is `CORREL`'s, a one-cell error before the sizes, then an error inside, then pairs. `SUMPRODUCT` beside
 * them as the engine computes it now. Answers: `test/excel/cases/binom-sumx-held.json`; `PREDICTED` was committed before
 * Excel ran the set (D!A7 is `1/0`, D!A8 `NA()`, D!B1:B8 numbers, D!E1:E8 numbers with `1/0` in E3).
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { n, e } from '../vals.js'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // under 171 trials as now, then past the factorials' range: a density, a range, the cumulative from 0
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(60,0.75,48)', n(0.083974967429047462)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(170,0.5,85)', n(0.06110502640760504)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(171,0.5,85)', n(0.060749764626165492)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(200,0.5,100)', n(0.056348479009256422)],
  ['BINOM.DIST', 'BINOM.DIST(100,200,0.5,FALSE)', n(0.056348479009256422)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(200,0.5,90,110)', n(0.86263334913136924)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(200.7,0.5,100.9)', n(0.056348479009256422)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(1000,0.5,500)', n(0.025225018178360804)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(1000,0.3,0,300)', n(0.51559351981412005)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(1030,0.5,515)', n(0.024855129936574471)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(2000,0.5,1000)', n(0.017839011145854337)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(10000,0.5,4900,5100)', n(0.95557420095391965)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(100000,0.01,1000)', n(0.012678161323544598)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(1000,0.5,0,1000)', n(1)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(1000,0.001,0)', n(0.36769542477096406)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(200,0.5,0)', n(6.223015277861121e-61)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(2000,0.5,0)', n(0)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(300,0.2,250,300)', n(8.4585703959526419e-123)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(1000000,0.5,499000,501000)', n(0.95460766409065612)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(1000000,0.5,500000)', n(0.0007978843613317499)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(2147483646,0.5,1073741823)', n(0.000017217699699243083)],
  ['BINOM.DIST.RANGE', 'BINOM.DIST.RANGE(2147483647,0.5,1)', e('NUM')],
  // an error in one cell against a range of another size, either side, each function
  ['SUMX2MY2', 'SUMX2MY2(D!A7,D!B1:B2)', e('DIV0')],
  ['SUMX2MY2', 'SUMX2MY2(D!B1:B2,D!A7)', e('DIV0')],
  ['SUMXMY2', 'SUMXMY2(D!A7,D!B1:B2)', e('DIV0')],
  ['SUMXMY2', 'SUMXMY2(D!B1:B3,D!A8)', e('NA')],
  ['SUMX2PY2', 'SUMX2PY2(D!A8,D!B1:B2)', e('NA')],
  ['SUMX2PY2', 'SUMX2PY2(D!B1:B2,D!A7)', e('DIV0')],
  // the same cell as a range of one, an array of one, a value computed, a reference IF gives
  ['SUMX2MY2', 'SUMX2MY2(D!A7:A7,D!B1:B2)', e('DIV0')],
  ['SUMX2MY2', 'SUMX2MY2({#DIV/0!},{1,2})', e('DIV0')],
  ['SUMX2MY2', 'SUMX2MY2(1/0,D!B1:B2)', e('DIV0')],
  ['SUMX2MY2', 'SUMX2MY2(IF(TRUE,D!A7),D!B1:B2)', e('DIV0')],
  ['SUMX2MY2', 'SUMX2MY2({1,2,3},D!A7)', e('DIV0')],
  ['SUMX2MY2', 'SUMX2MY2(D!A7,{1,2,3})', e('DIV0')],
  // both one cell, both errors; an error inside a range, sizes alike and not; a one-cell text and blank against a range
  ['SUMX2MY2', 'SUMX2MY2(D!A7,D!B1)', e('DIV0')],
  ['SUMX2MY2', 'SUMX2MY2(D!A7,D!A8)', e('DIV0')],
  ['SUMX2MY2', 'SUMX2MY2(D!A8,D!A7)', e('NA')],
  ['SUMX2MY2', 'SUMX2MY2(D!E1:E8,D!B1:B7)', e('NA')],
  ['SUMX2MY2', 'SUMX2MY2(D!E1:E8,D!B1:B8)', e('DIV0')],
  ['SUMX2MY2', 'SUMX2MY2(D!A3,D!B1:B2)', e('NA')],
  ['SUMX2MY2', 'SUMX2MY2(D!A1,D!B1:B2)', e('NA')],
  // the pairs of the statistics and SUMPRODUCT beside them
  ['CORREL', 'CORREL(D!A7,D!B1:B2)', e('DIV0')],
  ['PEARSON', 'PEARSON(D!B1:B2,D!A7)', e('DIV0')],
  ['SUMPRODUCT', 'SUMPRODUCT(D!A7,D!B1:B2)', e('VALUE')],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)

/**
 * The predictions Excel did not confirm (the Mac and Windows alike): an array of one error is no cell, so the sizes come
 * first (#N/A); a blank cell against a range is #VALUE!; SUMPRODUCT answers the cell's error before its sizes. The set
 * `sumx-cell-held` measured the rules these showed.
 */
export const MISSED: readonly string[] = ['SUMX2MY2#4', 'SUMX2MY2#15', 'SUMPRODUCT#1']
