/**
 * Held out after `near-zero-held`, which showed two rules: a formula's last + or - is 0 where its answer is below 8
 * units in the last place of the operands (2^(e-49) for an operand in [2^e, 2^(e+1)), at 1, 1.9, 9.5, 1E+15 and
 * 1E-200 alike), and SUM(0.1,0.2,-0.3) is 0 too. The rules predicted: the exponent is the larger operand's, which the
 * cases at the edge of a binade tell (`1-(1-k*2^-53)` and the other way round, read from cells of sheet Q times 2^53);
 * SUM adds its numbers in turn by the same rule, its arguments, ranges and arrays alike (a small number after the
 * cancellation is kept: `SUM(0.1,0.2,-0.3,1E-20)` is 1E-20), and so do SUBTOTAL, AGGREGATE, SUMIF, SUMIFS, AVERAGE and
 * SUMPRODUCT, while MMULT and NPV keep the double. Each answer is magnified where it is read (times 2^60, or over the
 * unit), as the comparison's tolerance takes 5.55E-17 for 0. Q!B1:B3 hold 0.1, 0.2, -0.3.
 * Answers: `test/excel/cases/sum-zero-held.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Probe, SheetCells } from '../probes.js'
import { heldGroups, xl } from './arrays/common.js'
import { n } from '../vals.js'

/** The rule predicted: a sum or difference below 2^(e-49), e the larger operand's exponent, is 0. */
function near(a: number, b: number, r: number): number {
  const e = Math.floor(Math.log2(Math.max(Math.abs(a), Math.abs(b))))
  return Math.abs(r) < 2 ** (e - 49) ? 0 : r
}

/** SUM as predicted: the numbers added in turn, each addition by the rule. */
const sum = (...xs: number[]) => xs.reduce((s, x) => near(s, x, s + x), 0)

const M = 2 ** 60
const K = [7, 8, 12, 15, 16, 17]

/** The binade's edge in cells of Q: k units of 2^-53 below 1 against 1, either way round. */
const EDGE: [string, number][] = [...K.map((k): [string, number] => [`1-(1-${k}*2^-53)`, near(1, 1 - k * 2 ** -53, k * 2 ** -53)]), ...K.map((k): [string, number] => [`(1-${k}*2^-53)-1`, near(1 - k * 2 ** -53, 1, -k * 2 ** -53)])]

export const SHEETS: SheetCells = {
  Q: { B1: 0.1, B2: 0.2, B3: -0.3, ...Object.fromEntries(EDGE.map(([f], i) => [`A${i + 1}`, { formula: xl(f) }])) },
}

/** Group, function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, string, number][] = [
  ...EDGE.map(([, v], i): [string, string, string, number] => ['edge', 'operators', `Q!A${i + 1}/2^-53`, v / 2 ** -53]),
  // SUM: arguments in other orders, a small number after the cancellation, the threshold, inside a product
  ['sum', 'SUM', 'SUM(0.1,0.2,-0.3)*2^60', sum(0.1, 0.2, -0.3) * M],
  ['sum', 'SUM', 'SUM(-0.3,0.1,0.2)*2^60', sum(-0.3, 0.1, 0.2) * M],
  ['sum', 'SUM', 'SUM(0.1,-0.3,0.2)*2^60', sum(0.1, -0.3, 0.2) * M],
  ['sum', 'SUM', 'SUM(0.1,0.2,-0.3,1E-20)/1E-20', sum(0.1, 0.2, -0.3, 1e-20) / 1e-20],
  ['sum', 'SUM', 'SUM(1E-20,0.1,0.2,-0.3)*2^60', sum(1e-20, 0.1, 0.2, -0.3) * M],
  ['sum', 'SUM', 'SUM(1,7*2^-52,-1)/2^-52', sum(1, 7 * 2 ** -52, -1) / 2 ** -52],
  ['sum', 'SUM', 'SUM(1,8*2^-52,-1)/2^-52', sum(1, 8 * 2 ** -52, -1) / 2 ** -52],
  ['sum', 'SUM', 'SUM(9.5,7*2^-49,-9.5)/2^-49', sum(9.5, 7 * 2 ** -49, -9.5) / 2 ** -49],
  ['sum', 'SUM', 'SUM(9.5,8*2^-49,-9.5)/2^-49', sum(9.5, 8 * 2 ** -49, -9.5) / 2 ** -49],
  ['sum', 'SUM', 'SUM(0.1+0.2,-0.3)*2^60', sum(0.1 + 0.2, -0.3) * M],
  // SUM over a range, an array, a range and a number
  ['sum', 'SUM', 'SUM(Q!B1:B3)*2^60', sum(0.1, 0.2, -0.3) * M],
  ['sum', 'SUM', 'SUM({0.1,0.2,-0.3})*2^60', sum(0.1, 0.2, -0.3) * M],
  ['sum', 'SUM', 'SUM(Q!B1:B2,-0.3)*2^60', sum(0.1, 0.2, -0.3) * M],
  // the functions that sum
  ['SUBTOTAL', 'SUBTOTAL', 'SUBTOTAL(9,Q!B1:B3)*2^60', sum(0.1, 0.2, -0.3) * M],
  ['AGGREGATE', 'AGGREGATE', 'AGGREGATE(9,0,Q!B1:B3)*2^60', sum(0.1, 0.2, -0.3) * M],
  ['SUMIF', 'SUMIF', 'SUMIF(Q!B1:B3,"<>0")*2^60', sum(0.1, 0.2, -0.3) * M],
  ['SUMIFS', 'SUMIFS', 'SUMIFS(Q!B1:B3,Q!B1:B3,"<>0")*2^60', sum(0.1, 0.2, -0.3) * M],
  ['AVERAGE', 'AVERAGE', 'AVERAGE(0.1,0.2,-0.3)*2^60', (sum(0.1, 0.2, -0.3) / 3) * M],
  ['AVERAGE', 'AVERAGE', 'AVERAGE(Q!B1:B3)*2^60', (sum(0.1, 0.2, -0.3) / 3) * M],
  ['SUMPRODUCT', 'SUMPRODUCT', 'SUMPRODUCT({0.1,0.2,-0.3},{1,1,1})*2^60', sum(0.1, 0.2, -0.3) * M],
  ['MMULT', 'MMULT', 'MMULT({1,1,1},{0.1;0.2;-0.3})*2^60', (0.1 + 0.2 - 0.3) * M],
  ['NPV', 'NPV', 'NPV(0,0.1,0.2,-0.3)*2^60', (0.1 + 0.2 - 0.3) * M],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldGroups(CASES.map(([group, fn, f, v]) => [group, fn, f, n(v)]))

/**
 * The predictions Excel did not confirm (the Mac and Windows alike): the exponent is the left operand's, not the larger
 * one's (`(1-8*2^-53)-1` keeps its 8 units, the bound 2^-50 of 1-8*2^-53); SUM takes the rule at its last addition
 * alone (`SUM(0.1,0.2,-0.3,1E-20)` is 5.55E-17 + 1E-20); SUMIF, SUMIFS and SUMPRODUCT keep the double.
 */
export const MISSED: readonly string[] = ['edge#8', 'edge#9', 'edge#10', 'sum#4', 'SUMIF#1', 'SUMIFS#1', 'SUMPRODUCT#1']
