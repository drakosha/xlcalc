/**
 * ACCRINT on the cases of IronCalc's tests (`xlsx/tests/calc_tests/FINANCIAL`, files saved by Excel) that the engine
 * answered otherwise, one or two for each kind: every basis, calc_method TRUE and FALSE, a settlement before the
 * period that ends at first_interest (a negative answer), a fractional frequency, the arguments read from cells. Then
 * the rule around them: coupon dates stepped one from the other (a day moved by a short month stays moved), a
 * first_interest on a month's last day, a settlement on a coupon date, an issue on one or after first_interest, the
 * issue's own period on each basis. Answers: `test/excel/cases/accrint-ironcalc-held.json`; `PREDICTED`, the engine's
 * answers by the new rule, was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { n } from '../vals.js'

/** Ac: the arguments of three rows of the IronCalc files (issue, first_interest, settlement, rate, par, frequency, basis, calc_method). */
export const SHEETS: SheetCells = {
  Ac: {
    A1: { formula: 'DATE(2017,1,1)' }, B1: { formula: 'DATE(2017,12,1)' }, C1: { formula: 'DATE(2017, 4,1)' }, D1: 0.33, E1: 3000, F1: 4, G1: 1, H1: false,
    A2: { formula: 'DATE(2020,2,3)' }, B2: { formula: 'DATE(2023,12,1)' }, C2: { formula: 'DATE(2023, 4,12)' }, D2: 0.33, E2: 3000, F2: 4.3, G2: 1, H2: false,
    A3: 44166, B3: 44531, C3: 44301, D3: 0.025, E3: 5000, F3: 4, G3: 1, H3: false,
  },
}

const row = (r: number, args: number) => `ACCRINT(${'ABCDEFGH'.slice(0, args).split('').map((c) => `Ac!${c}${r}`).join(',')})`

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // IronCalc's cases: an issue periods before the settlement, by basis
  ['ACCRINT', 'ACCRINT(DATE(2017,1,1), DATE(2017,12,1), DATE(2017,4,1), 0.33, 3000, 4, 1)', n(241.123626373626)],
  ['ACCRINT', 'ACCRINT(DATE(2017,1,1), DATE(2017,12,1), DATE(2017,4,1), 0.33, 3000, 4, 2)', n(236.5)],
  ['ACCRINT', 'ACCRINT(DATE(2017,1,1), DATE(2017,12,1), DATE(2017,4,1), 0.33, 3000, 4, 3)', n(240.041095890411)],
  ['ACCRINT', 'ACCRINT(DATE(2016,6,1), DATE(2017,12,1), DATE(2016,9,1), 0.33, 3000, 2, 2)', n(239.25)],
  ['ACCRINT', 'ACCRINT(DATE(2016,6,1), DATE(2017,12,1), DATE(2016,9,1), 0.33, 3000, 2, 1)', n(251.55737704918)],
  ['ACCRINT', 'ACCRINT(DATE(2007,3,1), DATE(2008,8,31), DATE(2008,5,1), 0.1, 1000, 2, 0, FALSE)', n(66.9444444444445)],
  // the arguments in cells
  ['ACCRINT', row(1, 8), n(-253.876373626374)],
  ['ACCRINT', row(1, 7), n(241.123626373626)],
  ['ACCRINT', row(2, 8), n(-312.774725274725)],
  ['ACCRINT', row(3, 8), n(-47.7335164835165)],
  // calc_method FALSE, the default basis, frequency 1, a fractional frequency, a settlement on first_interest
  ['ACCRINT', 'ACCRINT(39508,39691,39569,0.1,1000,4)', n(16.9444444444444)],
  ['ACCRINT', 'ACCRINT(42736,43070,42826,0.33,3000,2,1,FALSE)', n(245.686813186813)],
  ['ACCRINT', 'ACCRINT(43864,45261,45028,0.33,3000,1,1,FALSE)', n(1174.91264316191)],
  ['ACCRINT', 'ACCRINT(43864,45261,45028,0.33,3000,2,4,FALSE)', n(189.75)],
  ['ACCRINT', 'ACCRINT(43864,45261,45028,0.33,3000,4.7,3)', n(3153.08219178082)],
  ['ACCRINT', 'ACCRINT(44166,44531,44531,0.1,1000,2,2,FALSE)', n(50.8333333333333)],
  ['ACCRINT', 'ACCRINT(44191,44531,44256,0.025,5000,2,1,TRUE)', n(22.4940701375128)],
  ['ACCRINT', 'ACCRINT(44804,45351,45047,0.075,100,2,4,TRUE)', n(5.02083333333333)],
  ['ACCRINT', 'ACCRINT(43435,44531,43983,0.1,1000,1,0,FALSE)', n(-50)],
  ['ACCRINT', 'ACCRINT(44135,44408,44242,0.025,5000,4,1,TRUE)', n(37.3641304347826)],
  ['ACCRINT', 'ACCRINT(43845,44166,43922,0.05,1000,4,0,FALSE)', n(-14.4444444444444)],
  ['ACCRINT', 'ACCRINT(43800,44531,44440,0.1,1000,2,3,FALSE)', n(25.2054794520548)],
  // coupon dates stepped from first_interest one from the other: 30 August, then 28 February, then 28 August
  ['ACCRINT', 'ACCRINT(DATE(2024,10,1),DATE(2024,8,30),DATE(2025,9,15),0.05,1000,2,0)', n(47.5534188034188)],
  ['ACCRINT', 'ACCRINT(DATE(2024,3,15),DATE(2024,8,30),DATE(2025,9,15),0.05,1000,2,1)', n(75.2612141956404)],
  // and back: 29 February, 29 August, 28 February, 28 August
  ['ACCRINT', 'ACCRINT(DATE(2022,9,10),DATE(2024,8,30),DATE(2024,5,1),0.05,1000,2,1)', n(81.7036410073652)],
  ['ACCRINT', 'ACCRINT(DATE(2022,9,10),DATE(2024,8,30),DATE(2024,5,1),0.05,1000,2,0)', n(81.5491452991453)],
  // a first_interest on a month's last day: 30 April, then 31 October; 28 February of a leap year is not one
  ['ACCRINT', 'ACCRINT(DATE(2024,1,10),DATE(2024,4,30),DATE(2024,11,15),0.05,1000,2,1)', n(42.3076923076923)],
  ['ACCRINT', 'ACCRINT(DATE(2022,6,1),DATE(2024,2,28),DATE(2025,6,1),0.05,1000,1,1)', n(150)],
  ['ACCRINT', 'ACCRINT(DATE(2022,6,1),DATE(2024,2,29),DATE(2025,6,1),0.05,1000,1,1)', n(149.96519200539)],
  // a settlement on a coupon date after first_interest, an issue on a coupon date, an issue after first_interest
  ['ACCRINT', 'ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2025,7,1),0.05,1000,2,0)', n(75)],
  ['ACCRINT', 'ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2025,7,1),0.05,1000,2,1)', n(74.8626373626374)],
  ['ACCRINT', 'ACCRINT(DATE(2023,7,1),DATE(2024,7,1),DATE(2024,3,15),0.05,1000,2,1)', n(35.1648351648352)],
  ['ACCRINT', 'ACCRINT(DATE(2024,9,1),DATE(2024,7,1),DATE(2025,3,1),0.05,1000,2,0)', n(25)],
  ['ACCRINT', 'ACCRINT(DATE(2024,9,1),DATE(2024,7,1),DATE(2025,3,1),0.05,1000,2,0,FALSE)', n(25)],
  // calc_method FALSE with the settlement after first_interest: from the period that ends there, the issue's part before it
  ['ACCRINT', 'ACCRINT(DATE(2023,3,15),DATE(2024,1,31),DATE(2024,9,10),0.06,1000,4,0,FALSE)', n(59.1666666666667)],
  ['ACCRINT', 'ACCRINT(DATE(2023,3,15),DATE(2024,1,31),DATE(2024,9,10),0.06,1000,4,4,FALSE)', n(59.1666666666667)],
  ['ACCRINT', 'ACCRINT(DATE(2023,3,15),DATE(2024,1,31),DATE(2024,9,10),0.06,1000,4,3,TRUE)', n(89.3013698630137)],
  ['ACCRINT', 'ACCRINT(DATE(2024,2,1),DATE(2024,7,1),DATE(2025,3,1),0.05,1000,2,1,FALSE)', n(54.1208791208791)],
  ['ACCRINT', 'ACCRINT(DATE(2022,1,1),DATE(2024,6,30),DATE(2023,3,1),0.05,1000,1,0,FALSE)', n(8.33333333333333)],
  // the issue's own period across February on bases 2 and 0, and ending on a 31st on basis 4
  ['ACCRINT', 'ACCRINT(DATE(2023,9,15),DATE(2024,8,31),DATE(2024,10,15),0.05,1000,2,2)', n(54.4444444444444)],
  ['ACCRINT', 'ACCRINT(DATE(2023,9,15),DATE(2024,8,31),DATE(2024,10,15),0.05,1000,2,0)', n(54.0277777777778)],
  ['ACCRINT', 'ACCRINT(DATE(2023,10,31),DATE(2024,8,31),DATE(2024,9,30),0.05,1000,4,4)', n(45.8333333333333)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
