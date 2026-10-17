/**
 * Math and statistics on the cases of IronCalc's tests (`xlsx/tests/calc_tests`, files saved by Excel) that the
 * engine answered otherwise: GCD and LCM of blank cells alone, CSCH and IMCSCH of a tiny argument, LOGNORM.INV of
 * LOGNORM.DIST with σ = 1E12 (where one unit in the last place of the probability moves the answer by 1E-4),
 * POISSON.DIST of a hair below 0, PROB over text. Then the rules around them. Answers:
 * `test/excel/cases/math-ironcalc-held.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { e, n, s } from '../vals.js'

/** Mt: blanks in A1:C2, text pairs in D1:E2 as in IronCalc's PROB.xlsx, text in F, a text pair among numbers in G1:H3. */
export const SHEETS: SheetCells = {
  Mt: { D1: 'Hi', D2: 'ho', E1: 'hi', E2: 'ho', F1: 'a', F2: 'b', G1: 1, G2: 'x', G3: 2, H1: 0.5, H2: 0.2, H3: 0.5 },
}

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // IronCalc's cases
  ['GCD', 'GCD(Mt!A1,Mt!B1,Mt!C1)', e('VALUE')],
  ['LCM', 'LCM(Mt!A1,Mt!B1,Mt!C1)', e('VALUE')],
  ['CSCH', 'CSCH(1E-12)', n(1e12)],
  ['CSCH', 'CSCH(-1E-12)', n(-1e12)],
  ['IMREAL', 'IMREAL(IMCSCH(COMPLEX(1E-16,0)))', n(1e16)],
  ['IMAGINARY', 'IMAGINARY(IMCSCH(COMPLEX(1E-20,0)))', n(0)],
  ['LOGNORM.INV', 'LOGNORM.INV(LOGNORM.DIST(4,3.5,1E12,TRUE),3.5,1E12)', n(3.99976588093441)],
  ['POISSON.DIST', 'POISSON.DIST(-0.000001,5,TRUE)', e('NUM')],
  ['PROB', 'PROB(Mt!D1:D2,Mt!E1:E2,1)', e('DIV0')],
  // GCD and LCM: one blank, a blank beside a number, a blank range, text
  ['GCD', 'GCD(Mt!A1)', e('VALUE')],
  ['GCD', 'GCD(Mt!A1,6)', n(6)],
  ['LCM', 'LCM(Mt!A1,4)', n(4)],
  ['GCD', 'GCD(Mt!A1:A2)', n(0)],
  ['GCD', 'GCD(Mt!F1:F2)', e('VALUE')],
  // CSCH and IMCSCH at other sizes
  ['CSCH', 'CSCH(1E-300)', n(1e300)],
  ['CSCH', 'CSCH(0.5)', n(1.91903475133494)],
  ['IMCSCH', 'IMCSCH("1E-300")', s('1E+300')],
  ['IMAGINARY', 'IMAGINARY(IMCSCH(COMPLEX(0,1E-16)))', n(-1e16)],
  ['IMCSCH', 'IMCSCH("0.5+0.5i")', s('0.912074264038811-1.07822969465402i')],
  ['IMCSCH', 'IMCSCH(0)', e('NUM')],
  // the normal distribution next to 0, and further out
  ['LOGNORM.INV', 'LOGNORM.INV(LOGNORM.DIST(4,3.5,1E8,TRUE),3.5,1E8)', n(3.99999997301232)],
  ['NORM.S.DIST', 'NORM.S.DIST(1E-13,TRUE)-0.5', n(3.98942280401433e-14)],
  ['NORM.S.DIST', 'NORM.S.DIST(0.2,TRUE)', n(0.579259709439103)],
  ['NORM.S.INV', 'NORM.S.INV(0.5-2^-40)', n(-2.27966904567376e-12)],
  // POISSON.DIST: 0 and a fraction truncated, a hair below 0
  ['POISSON.DIST', 'POISSON.DIST(-0,5,TRUE)', n(0.00673794699908547)],
  ['POISSON.DIST', 'POISSON.DIST(0.9,5,FALSE)', n(0.00673794699908547)],
  ['POISSON.DIST', 'POISSON.DIST(-1E-300,5,TRUE)', e('NUM')],
  // PROB passes over pairs that are not two numbers
  ['PROB', 'PROB(Mt!G1:G3,Mt!H1:H3,1)', n(0.5)],
  ['PROB', 'PROB({"a","b"},{0.5,0.5},1)', e('DIV0')],
  ['PROB', 'PROB({TRUE,1},{0.5,1},1)', n(1)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)

/**
 * All 30 as predicted within the comparison's tolerance, the Mac and Windows alike. To the last digit NORM.S.DIST of
 * 1E-13 is not what the prediction said (erf's series) but what erfc gives (3.9968028886505635E-14 above 0.5), while
 * LOGNORM.DIST next to 0.5 is the series' value: the engine takes the series in LOGNORM.DIST alone.
 */
export const MISSED: readonly string[] = []
