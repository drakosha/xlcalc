/**
 * Statistics on edge data, held out: the cases of `test/statistical-stability.test.ts` (a small spread around a large
 * mean, scales 1E150 and 1E-200, a variance that overflows, a constant near the largest double, covariance across
 * scales, STEYX near a line and at a strong correlation, a small intercept after cancellation, the database
 * statistics) with the moment and regression functions beside them on the same data. The engine computes these with
 * compensated and exact sums; whether Excel's answers are the exact ones was not measured before. Answers:
 * `test/excel/cases/stat-edges-held.json`; `PREDICTED`, the engine's answers, was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'

/** The large-mean data: 10,000 rows of 1E12 + 0, 0.01, 0.02 in turn (A) and their step 0, 1, 2 (B). */
const ROWS = 10000
const big: Record<string, number | string> = { A1: 'big', B1: 'k' }
for (let i = 0; i < ROWS; i++) {
  big[`A${i + 2}`] = 1e12 + (i % 3) / 100
  big[`B${i + 2}`] = i % 3
}

/** A column from row 1: its header, then its values. */
const column = (c: string, header: string, ...xs: (number | { formula: string })[]) => Object.fromEntries([[`${c}1`, header], ...xs.map((x, i) => [`${c}${i + 2}`, x])])

/** ±1E308 by a formula: Excel takes a file holding 1E+308 as a value for a damaged one (above 9.99999999999999E+307). */
const MAX = { formula: '1E+307*10' }, MIN = { formula: '-1E+307*10' }

/**
 * Sheet Edge: each data set under a header, so the database functions take it too; Q1:U2 the criteria (a header over
 * a blank: every record).
 */
export const SHEETS: SheetCells = {
  Edge: {
    ...big,
    ...column('C', 'h', 1e150, 2e150),
    ...column('D', 't', 1e-200, 2e-200),
    ...column('E', 'u', 1, 2),
    ...column('F', 'o', 1e300, -1e300),
    ...column('G', 'm', MAX, MAX, MAX),
    ...column('H', 'cx', MAX, MIN, 1e-200),
    ...column('I', 'cy', 0, 0, MAX),
    ...column('J', 'sx', 1, 2, 3),
    ...column('K', 'sy', 1, 2.00000001, 3),
    ...column('L', 'ix', 0.25, 0.5, 0.75),
    ...column('M', 'iy', 0.5, 1.00000001, 1.5),
    ...column('N', 'rx', 1.218261304197592e-10, -1.25634210083295e-11, 1.4598980988767263e-10),
    ...column('O', 'ry', 6.150589335415155e-11, 1.9303787081095429e-10, 3.8973379524949484e-11),
    Q1: 'big', R1: 't', S1: 'o', T1: 'm', U1: 'h',
  },
}

const n = (v: number): Val => ({ t: 'n', v })
const e = (v: Extract<Val, { t: 'e' }>['v']): Val => ({ t: 'e', v })

const BIG = 'Edge!A2:A10001', STEP = 'Edge!B2:B10001'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // the data as Excel read them: 1E12 + 82/8192, the extremes, the digits past the 15th
  ['data', '(Edge!A3-1E12)*8192', n(82)],
  ['data', 'Edge!G2/1E+307', n(10)],
  ['data', 'Edge!D2/1E-200', n(1)],
  ['data', '(Edge!N2-1.21826130419759E-10)*1E25', n(2.06795153138257)],
  ['data', '(Edge!O3-1.93037870810954E-10)*1E25', n(2.84343335565103)],
  ['data', '(Edge!K3-2)*1E8', n(0.999999993922529)],
  // a small spread around a large mean: 10,000 values
  ['AVERAGE', `AVERAGE(${BIG})`, n(1000000000000.01)],
  ['AVERAGE', 'AVERAGE(Edge!A:A)', n(1000000000000.01)],
  ['VAR.P', `VAR.P(${BIG})`, n(0.0000668002774232626)],
  ['VAR.P', 'VAR.P(Edge!A:A)', n(0.0000668002774232626)],
  ['VAR.S', `VAR.S(${BIG})`, n(0.0000668069581190745)],
  ['STDEV.P', `STDEV.P(${BIG})`, n(0.00817314366833611)],
  ['STDEV.P', 'STDEV.P(Edge!A:A)', n(0.00817314366833611)],
  ['STDEV.S', `STDEV.S(${BIG})`, n(0.00817355235617137)],
  ['DEVSQ', `DEVSQ(${BIG})`, n(0.668002774232626)],
  ['DEVSQ', 'DEVSQ(Edge!A:A)', n(0.668002774232626)],
  ['COVARIANCE.P', `COVARIANCE.P(${BIG},${STEP})`, n(0.00667351064208984)],
  ['COVARIANCE.S', `COVARIANCE.S(${BIG},${STEP})`, n(0.00667417805989583)],
  ['CORREL', `CORREL(${BIG},${STEP})`, n(1)],
  ['PEARSON', `PEARSON(${BIG},${STEP})`, n(1)],
  ['RSQ', `RSQ(${BIG},${STEP})`, n(1)],
  ['SLOPE', `SLOPE(${BIG},${STEP})`, n(0.010009765625)],
  ['SLOPE', `SLOPE(${STEP},${BIG})`, n(99.9024390243902)],
  ['INTERCEPT', `INTERCEPT(${BIG},${STEP})`, n(1000000000000)],
  ['FORECAST.LINEAR', `FORECAST.LINEAR(3,${BIG},${STEP})`, n(1000000000000.03)],
  ['STEYX', `STEYX(${BIG},${STEP})`, n(0)],
  ['LINEST', `INDEX(LINEST(${BIG},${STEP}),1,1)`, n(0.0101547267909801)],
  ['LINEST', `INDEX(LINEST(${BIG},${STEP}),1,2)`, n(999999999999.99)],
  ['DAVERAGE', 'DAVERAGE(Edge!A1:A10001,1,Edge!Q1:Q2)', n(1000000000000.01)],
  ['DVAR', 'DVAR(Edge!A1:A10001,1,Edge!Q1:Q2)', n(0.0000668069581190745)],
  ['DVARP', 'DVARP(Edge!A1:A10001,1,Edge!Q1:Q2)', n(0.0000668002774232626)],
  ['DSTDEV', 'DSTDEV(Edge!A1:A10001,1,Edge!Q1:Q2)', n(0.00817355235617137)],
  ['DSTDEVP', 'DSTDEVP(Edge!A1:A10001,1,Edge!Q1:Q2)', n(0.00817314366833611)],
  // scale 1E150: 1E150 and 2E150 against 1 and 2
  ['AVERAGE', 'AVERAGE(Edge!C2:C3)', n(1.5e+150)],
  ['VAR.P', 'VAR.P(Edge!C2:C3)', n(2.5e+299)],
  ['VAR.S', 'VAR.S(Edge!C2:C3)', n(5e+299)],
  ['STDEV.P', 'STDEV.P(Edge!C2:C3)', n(5e+149)],
  ['STDEV.S', 'STDEV.S(Edge!C2:C3)', n(7.07106781186548e+149)],
  ['DEVSQ', 'DEVSQ(Edge!C2:C3)', n(5e+299)],
  ['COVARIANCE.S', 'COVARIANCE.S(Edge!C2:C3,Edge!E2:E3)', n(5e+149)],
  ['CORREL', 'CORREL(Edge!C2:C3,Edge!E2:E3)', n(1)],
  ['PEARSON', 'PEARSON(Edge!C2:C3,Edge!E2:E3)', n(1)],
  ['RSQ', 'RSQ(Edge!C2:C3,Edge!E2:E3)', n(1)],
  ['SLOPE', 'SLOPE(Edge!C2:C3,Edge!E2:E3)', n(1e+150)],
  ['SLOPE', 'SLOPE(Edge!E2:E3,Edge!C2:C3)', n(1e-150)],
  ['FORECAST.LINEAR', 'FORECAST.LINEAR(3,Edge!C2:C3,Edge!E2:E3)', n(3e+150)],
  ['DSTDEV', 'DSTDEV(Edge!C1:C3,1,Edge!U1:U2)', n(7.07106781186548e+149)],
  // scale 1E-200: the squares underflow
  ['AVERAGE', 'AVERAGE(Edge!D2:D3)', n(1.5e-200)],
  ['VAR.P', 'VAR.P(Edge!D2:D3)', n(0)],
  ['VAR.S', 'VAR.S(Edge!D2:D3)', n(0)],
  ['STDEV.P', 'STDEV.P(Edge!D2:D3)', n(5e-201)],
  ['STDEV.S', 'STDEV.S(Edge!D2:D3)', n(7.07106781186547e-201)],
  ['DEVSQ', 'DEVSQ(Edge!D2:D3)', n(0)],
  ['COVARIANCE.S', 'COVARIANCE.S(Edge!D2:D3,Edge!E2:E3)', n(5e-201)],
  ['CORREL', 'CORREL(Edge!D2:D3,Edge!E2:E3)', n(1)],
  ['PEARSON', 'PEARSON(Edge!D2:D3,Edge!E2:E3)', n(1)],
  ['RSQ', 'RSQ(Edge!D2:D3,Edge!E2:E3)', n(1)],
  ['SLOPE', 'SLOPE(Edge!D2:D3,Edge!E2:E3)', n(1e-200)],
  ['SLOPE', 'SLOPE(Edge!E2:E3,Edge!D2:D3)', n(1e+200)],
  ['FORECAST.LINEAR', 'FORECAST.LINEAR(3,Edge!D2:D3,Edge!E2:E3)', n(3e-200)],
  ['DAVERAGE', 'DAVERAGE(Edge!D1:D3,"t",Edge!R1:R2)', n(1.5e-200)],
  ['DVAR', 'DVAR(Edge!D1:D3,"t",Edge!R1:R2)', n(0)],
  ['DVARP', 'DVARP(Edge!D1:D3,"t",Edge!R1:R2)', n(0)],
  ['DSTDEV', 'DSTDEV(Edge!D1:D3,"t",Edge!R1:R2)', n(7.07106781186547e-201)],
  ['DSTDEVP', 'DSTDEVP(Edge!D1:D3,"t",Edge!R1:R2)', n(5e-201)],
  // 1E300 and -1E300: the variance itself overflows
  ['AVERAGE', 'AVERAGE(Edge!F2:F3)', n(0)],
  ['VAR.P', 'VAR.P(Edge!F2:F3)', e('NUM')],
  ['VAR.S', 'VAR.S(Edge!F2:F3)', e('NUM')],
  ['STDEV.P', 'STDEV.P(Edge!F2:F3)', e('NUM')],
  ['STDEV.S', 'STDEV.S(Edge!F2:F3)', e('NUM')],
  ['DEVSQ', 'DEVSQ(Edge!F2:F3)', e('NUM')],
  ['DSTDEV', 'DSTDEV(Edge!F1:F3,1,Edge!S1:S2)', e('NUM')],
  ['DVARP', 'DVARP(Edge!F1:F3,1,Edge!S1:S2)', e('NUM')],
  // a constant 1E308 three times: the sum overflows, the variance is 0
  ['AVERAGE', 'AVERAGE(Edge!G2:G4)', n(1e+308)],
  ['VAR.P', 'VAR.P(Edge!G2:G4)', n(0)],
  ['VAR.S', 'VAR.S(Edge!G2:G4)', n(0)],
  ['STDEV.P', 'STDEV.P(Edge!G2:G4)', n(0)],
  ['STDEV.S', 'STDEV.S(Edge!G2:G4)', n(0)],
  ['DEVSQ', 'DEVSQ(Edge!G2:G4)', n(0)],
  ['DAVERAGE', 'DAVERAGE(Edge!G1:G4,1,Edge!T1:T2)', n(1e+308)],
  ['DVAR', 'DVAR(Edge!G1:G4,1,Edge!T1:T2)', n(0)],
  // covariance across scales: x 1E308, -1E308, 1E-200 against y 0, 0, 1E308
  ['COVARIANCE.P', 'COVARIANCE.P(Edge!H2:H4,Edge!I2:I4)', n(2.22222222222222e+107)],
  ['COVARIANCE.S', 'COVARIANCE.S(Edge!H2:H4,Edge!I2:I4)', n(3.33333333333333e+107)],
  ['CORREL', 'CORREL(Edge!H2:H4,Edge!I2:I4)', n(0)],
  ['PEARSON', 'PEARSON(Edge!H2:H4,Edge!I2:I4)', n(0)],
  ['RSQ', 'RSQ(Edge!H2:H4,Edge!I2:I4)', n(0)],
  ['SLOPE', 'SLOPE(Edge!I2:I4,Edge!H2:H4)', n(0)],
  ['INTERCEPT', 'INTERCEPT(Edge!I2:I4,Edge!H2:H4)', n(3.33333333333333e+307)],
  ['FORECAST.LINEAR', 'FORECAST.LINEAR(1,Edge!I2:I4,Edge!H2:H4)', n(3.33333333333333e+307)],
  // STEYX near a straight line: y 1, 2.00000001, 3 on x 1, 2, 3
  ['STEYX', 'STEYX(Edge!K2:K4,Edge!J2:J4)', n(8.16496575965492e-9)],
  ['SLOPE', 'SLOPE(Edge!K2:K4,Edge!J2:J4)', n(1)],
  ['INTERCEPT', 'INTERCEPT(Edge!K2:K4,Edge!J2:J4)', n(3.3333333130751e-9)],
  ['CORREL', 'CORREL(Edge!K2:K4,Edge!J2:J4)', n(1)],
  ['PEARSON', 'PEARSON(Edge!K2:K4,Edge!J2:J4)', n(1)],
  ['RSQ', 'RSQ(Edge!K2:K4,Edge!J2:J4)', n(1)],
  ['COVARIANCE.S', 'COVARIANCE.S(Edge!K2:K4,Edge!J2:J4)', n(1)],
  ['FORECAST.LINEAR', 'FORECAST.LINEAR(4,Edge!K2:K4,Edge!J2:J4)', n(4.00000000333333)],
  ['LINEST', 'INDEX(LINEST(Edge!K2:K4,Edge!J2:J4,TRUE,TRUE),3,2)', n(8.16496537048295e-9)],
  // a small intercept after cancellation: y 0.5, 1.00000001, 1.5 on x 0.25, 0.5, 0.75
  ['INTERCEPT', 'INTERCEPT(Edge!M2:M4,Edge!L2:L4)', n(3.3333333130751e-9)],
  ['FORECAST', 'FORECAST(0,Edge!M2:M4,Edge!L2:L4)', n(3.3333333130751e-9)],
  ['FORECAST.LINEAR', 'FORECAST.LINEAR(0,Edge!M2:M4,Edge!L2:L4)', n(3.3333333130751e-9)],
  ['SLOPE', 'SLOPE(Edge!M2:M4,Edge!L2:L4)', n(2)],
  ['STEYX', 'STEYX(Edge!M2:M4,Edge!L2:L4)', n(8.16496575965492e-9)],
  ['LINEST', 'INDEX(LINEST(Edge!M2:M4,Edge!L2:L4),1,1)', n(2)],
  ['LINEST', 'INDEX(LINEST(Edge!M2:M4,Edge!L2:L4),1,2)', n(3.33333351641905e-9)],
  // STEYX at a strong correlation, values of 16 and 17 digits
  ['STEYX', 'STEYX(Edge!O2:O4,Edge!N2:N4)', n(7.17637025164989e-13)],
  ['CORREL', 'CORREL(Edge!O2:O4,Edge!N2:N4)', n(-0.999981405088139)],
  ['RSQ', 'RSQ(Edge!O2:O4,Edge!N2:N4)', n(0.999962810522049)],
  ['SLOPE', 'SLOPE(Edge!O2:O4,Edge!N2:N4)', n(-0.974073666813056)],
  ['INTERCEPT', 'INTERCEPT(Edge!O2:O4,Edge!N2:N4)', n(1.80717300375667e-10)],
  ['LINEST', 'INDEX(LINEST(Edge!O2:O4,Edge!N2:N4,TRUE,TRUE),3,2)', n(7.17637025164971e-13)],
]

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [fn, f] of CASES) add(fn, fn, f)
  return out
}

/** The predicted answers by case id. */
export const PREDICTED: Readonly<Record<string, Val>> = Object.fromEntries(buildProbes().map((p, i) => [p.id, CASES[i][2]]))

/**
 * The predictions Excel did not confirm, the Mac and Windows alike, all known differences (`KNOWN`): Excel computes the
 * moments in double precision, so the spread of the large-mean data is lost, the squares of 1E-200 underflow (CORREL
 * and SLOPE #DIV/0!), sums past 1.8E+308 are #NUM!, STEYX next to a line is 0; and the engine's LINEST lost digits
 * on the large mean that Excel keeps (LINEST#1), which it keeps now (`fit-offset-held`).
 */
export const MISSED: readonly string[] = [
  'VAR.P#1', 'VAR.P#2', 'VAR.S#1', 'STDEV.P#1', 'STDEV.P#2', 'STDEV.S#1', 'DEVSQ#1', 'DEVSQ#2', 'CORREL#1',
  'PEARSON#1', 'RSQ#1', 'SLOPE#2', 'STEYX#1', 'LINEST#1', 'DVAR#1', 'DVARP#1', 'DSTDEV#1', 'DSTDEVP#1', 'CORREL#3',
  'PEARSON#3', 'RSQ#3', 'SLOPE#6', 'AVERAGE#6', 'VAR.P#6', 'VAR.S#5', 'STDEV.P#6', 'STDEV.S#5', 'DEVSQ#6',
  'DAVERAGE#3', 'DVAR#3', 'COVARIANCE.P#2', 'COVARIANCE.S#4', 'CORREL#4', 'PEARSON#4', 'RSQ#4', 'SLOPE#7',
  'INTERCEPT#2', 'FORECAST.LINEAR#4', 'STEYX#2', 'STEYX#3',
]
