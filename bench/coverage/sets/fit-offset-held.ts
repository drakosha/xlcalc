/**
 * LINEST, LOGEST, TREND and GROWTH on data far from 0 against their spread, held out: the large-mean data of
 * `stat-edges-held` (10,000 values 1E12 + 0, 0.01, 0.02 on their steps 0, 1, 2) with a second variable and without the
 * constant, x far from 0 (1E10 + 0…11) for a line and a growth, and both far from 0 with two variables. The engine takes
 * y and each variable about their means before the decomposition when the model has a constant. Answers:
 * `test/excel/cases/fit-offset-held.json`; `PREDICTED`, the engine's answers, was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'
import { SHEETS as EDGE } from './stat-edges-held.js'

const ROWS = 10000
const K = 12
/** Small residuals, exact in binary at the data's scale. */
const NOISE = [0.5, -0.25, 0, 0.75, -0.5, 0.25, 0, -0.75, 0.5, -0.25, 0.25, -0.5]

/**
 * Sheet Fit: A2:B10001 the steps 0, 1, 2 and a second variable 0…6 beside the large-mean data of Edge!A; D2:D13 x =
 * 1E10 + 0…11, E2:E13 a line on it (3 + 2k and a residual), F2:F13 a growth (100·1.1^k to cents); H2:I13 x1 = 1E6 + k,
 * x2 = 1E6 + k² mod 7, J2:J13 y = 1E9 + k/2 − (k² mod 7)/4 and an eighth of the residual.
 */
const fit: Record<string, number | string> = { A1: 'k', B1: 'k7', D1: 'x', E1: 'line', F1: 'growth', H1: 'x1', I1: 'x2', J1: 'y' }
for (let i = 0; i < ROWS; i++) {
  fit[`A${i + 2}`] = i % 3
  fit[`B${i + 2}`] = i % 7
}
for (let k = 0; k < K; k++) {
  const r = k + 2
  fit[`D${r}`] = 1e10 + k
  fit[`E${r}`] = 3 + 2 * k + NOISE[k]
  fit[`F${r}`] = Math.round(100 * 1.1 ** k * 100) / 100
  fit[`H${r}`] = 1e6 + k
  fit[`I${r}`] = 1e6 + ((k * k) % 7)
  fit[`J${r}`] = 1e9 + k / 2 - ((k * k) % 7) / 4 + NOISE[k] / 8
}

export const SHEETS: SheetCells = { Edge: EDGE.Edge, Fit: fit }

const n = (v: number): Val => ({ t: 'n', v })

const BIG = 'Edge!A2:A10001', STEP = 'Edge!B2:B10001', TWO = 'Fit!A2:B10001'
const X = 'Fit!D2:D13', LINE = 'Fit!E2:E13', GROW = 'Fit!F2:F13', X2 = 'Fit!H2:I13', Y2 = 'Fit!J2:J13'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // the large-mean data on their steps, with a second variable 0…6 and without the constant: they lie on a line
  ['LINEST', `INDEX(LINEST(${BIG},${STEP}),1,2)-1E12`, n(0)],
  ['LINEST', `INDEX(LINEST(${BIG},${TWO}),1,1)`, n(-3.91358693204563e-18)],
  ['LINEST', `INDEX(LINEST(${BIG},${TWO}),1,2)`, n(0.0100097656250001)],
  ['LINEST', `INDEX(LINEST(${BIG},${TWO}),1,3)-1E12`, n(0)],
  ['LINEST', `INDEX(LINEST(${BIG},${STEP},FALSE),1,1)`, n(600000000000.001)],
  ['LINEST', `INDEX(LINEST(${BIG},${TWO},FALSE),1,1)`, n(157908869645.205)],
  ['LINEST', `INDEX(LINEST(${BIG},${TWO},FALSE),1,2)`, n(315801936557.539)],
  // its statistics: a perfect fit
  ['LINEST', `INDEX(LINEST(${BIG},${STEP},TRUE,TRUE),2,1)`, n(1.03602521372535e-17)],
  ['LINEST', `INDEX(LINEST(${BIG},${STEP},TRUE,TRUE),3,1)`, n(1)],
  ['LINEST', `INDEX(LINEST(${BIG},${STEP},TRUE,TRUE),3,2)`, n(8.45932185929357e-16)],
  ['LINEST', `INDEX(LINEST(${BIG},${STEP},TRUE,TRUE),5,1)`, n(0.668002774232626)],
  ['LINEST', `INDEX(LINEST(${BIG},${STEP},TRUE,TRUE),5,2)`, n(7.15458142938583e-27)],
  // TREND, LOGEST and GROWTH on them; less 1E12 the digits show
  ['TREND', `TREND(${BIG},${STEP},3)-1E12`, n(0.030029296875)],
  ['TREND', `TREND(${BIG},${TWO},{3,5})-1E12`, n(0.030029296875)],
  ['LOGEST', `INDEX(LOGEST(${BIG},${STEP}),1,1)`, n(1.00000000000001)],
  ['LOGEST', `INDEX(LOGEST(${BIG},${STEP}),1,2)`, n(999999999999.999)],
  ['GROWTH', `GROWTH(${BIG},${STEP},3)`, n(1000000000000.03)],
  // x = 1E10 + 0…11: a line, its statistics, its fitted values past and on the data
  ['LINEST', `INDEX(LINEST(${LINE},${X}),1,1)`, n(1.96328671328671)],
  ['LINEST', `INDEX(LINEST(${LINE},${X}),1,2)`, n(-19632867129.6652)],
  ['LINEST', `INDEX(LINEST(${LINE},${X},TRUE,TRUE),2,1)`, n(0.0390647234562074)],
  ['LINEST', `INDEX(LINEST(${LINE},${X},TRUE,TRUE),2,2)`, n(390647234.77693)],
  ['LINEST', `INDEX(LINEST(${LINE},${X},TRUE,TRUE),3,1)`, n(0.996056462173472)],
  ['LINEST', `INDEX(LINEST(${LINE},${X},TRUE,TRUE),5,2)`, n(2.18225524475524)],
  ['LINEST', `INDEX(LINEST(${LINE},${X},FALSE),1,1)`, n(1.39999999946396e-9)],
  ['TREND', `TREND(${LINE},${X},1E10+12)`, n(26.7613636363636)],
  ['TREND', `INDEX(TREND(${LINE},${X}),1)`, n(3.20192307692307)],
  // a growth on the same x
  ['LOGEST', `INDEX(LOGEST(${GROW},${X}),1,1)`, n(1.09999874875424)],
  ['LOGEST', `INDEX(LOGEST(${GROW},${X},TRUE,TRUE),2,1)`, n(8.6814400580925e-7)],
  ['LOGEST', `INDEX(LOGEST(${GROW},${X},TRUE,TRUE),3,1)`, n(0.999999999170311)],
  ['GROWTH', `GROWTH(${GROW},${X},1E10+12)`, n(313.839727343595)],
  ['GROWTH', `INDEX(GROWTH(${GROW},${X}),12)`, n(285.30916757771)],
  // x1 = 1E6 + k, x2 = 1E6 + k² mod 7, y near 1E9: two variables, with and without the constant
  ['LINEST', `INDEX(LINEST(${Y2},${X2}),1,1)`, n(-0.246865086587966)],
  ['LINEST', `INDEX(LINEST(${Y2},${X2}),1,2)`, n(0.49522449815383)],
  ['LINEST', `INDEX(LINEST(${Y2},${X2}),1,3)-1E9`, n(-248359.391309142)],
  ['LINEST', `INDEX(LINEST(${Y2},${X2},TRUE,TRUE),2,1)`, n(0.01295983958366)],
  ['LINEST', `INDEX(LINEST(${Y2},${X2},TRUE,TRUE),2,3)`, n(13225.2607522832)],
  ['LINEST', `INDEX(LINEST(${Y2},${X2},TRUE,TRUE),3,1)`, n(0.999015833822645)],
  ['LINEST', `INDEX(LINEST(${Y2},${X2},TRUE,TRUE),3,2)`, n(0.0613527920847263)],
  ['LINEST', `INDEX(LINEST(${Y2},${X2},TRUE,TRUE),4,1)`, n(4567.89854766669)],
  ['LINEST', `INDEX(LINEST(${Y2},${X2},FALSE,TRUE),1,1)`, n(902.715961044297)],
  ['LINEST', `INDEX(LINEST(${Y2},${X2},FALSE,TRUE),1,2)`, n(97.2817759693838)],
  ['LINEST', `INDEX(LINEST(${Y2},${X2},FALSE,TRUE),3,1)`, n(0.999999999998207)],
  ['TREND', `TREND(${Y2},${X2},{1000012,1000003})`, n(1000000005.22236)],
]

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [fn, f] of CASES) add(fn, fn, f)
  return out
}

/** The predicted answers by case id. */
export const PREDICTED: Readonly<Record<string, Val>> = Object.fromEntries(buildProbes().map((p, i) => [p.id, CASES[i][2]]))
