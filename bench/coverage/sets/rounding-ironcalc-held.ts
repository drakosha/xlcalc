/**
 * Rounding on the cases of IronCalc's tests (`xlsx/tests/calc_tests`, files saved by Excel) that the engine answered
 * otherwise: FLOOR of 0 (as a number, text, FALSE or a blank) to a significance of 0, FLOOR.MATH and CEILING.MATH of
 * negative numbers with the mode 0.01, ROUND of a difference just below a half. Then the rules around them: FLOOR of
 * other numbers to 0, modes small, negative, text and boolean, positive numbers, ROUND just below a half at other
 * magnitudes. Answers: `test/excel/cases/rounding-ironcalc-held.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { e, n } from '../vals.js'

/** Fl: 0 as a number, as text and as FALSE, a blank (A4), and IronCalc's difference just below 0.00625 (B1). */
export const SHEETS: SheetCells = {
  Fl: { A1: 0, A2: { formula: '"0"' }, A3: false, B1: { formula: '1.05*(0.0284+0.0046)-0.0284' } },
}

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // IronCalc's cases: 0 to a significance of 0, in any form
  ['FLOOR', 'FLOOR(0,0)', n(0)],
  ['FLOOR', 'FLOOR(Fl!A2,Fl!A3)', n(0)],
  ['FLOOR', 'FLOOR(Fl!A4,Fl!A4)', n(0)],
  ['FLOOR', 'FLOOR(Fl!A3,Fl!A2)', n(0)],
  // the mode 0.01 turns the rounding of negative numbers
  ['FLOOR.MATH', 'FLOOR.MATH(-7.1,1,0.01)', n(-7)],
  ['FLOOR.MATH', 'FLOOR.MATH(-10,0.21,0.01)', n(-9.87)],
  ['FLOOR.MATH', 'FLOOR.MATH(-7.9,-7,0.01)', n(-7)],
  ['CEILING.MATH', 'CEILING.MATH(-7.1,-1,0.01)', n(-8)],
  ['CEILING.MATH', 'CEILING.MATH(-7.9,0.21,0.01)', n(-7.98)],
  ['CEILING.MATH', 'CEILING.MATH(-10,7,0.01)', n(-14)],
  ['ROUND', 'ROUND(1.05*(0.0284+0.0046)-0.0284,4)', n(0.0063)],
  ['ROUND', 'ROUND(Fl!B1,4)', n(0.0063)],
  // FLOOR of other numbers to 0, and of 0 to other significances
  ['FLOOR', 'FLOOR(5,0)', e('DIV0')],
  ['FLOOR', 'FLOOR(1E-300,0)', e('DIV0')],
  ['FLOOR', 'FLOOR(-5,0)', e('DIV0')],
  ['FLOOR', 'FLOOR(0,5)', n(0)],
  ['FLOOR', 'FLOOR(0,-5)', n(0)],
  ['CEILING', 'CEILING(0,0)', n(0)],
  // modes small, negative, as text and booleans, and positive numbers
  ['FLOOR.MATH', 'FLOOR.MATH(-7.1,1,-0.01)', n(-7)],
  ['FLOOR.MATH', 'FLOOR.MATH(-7.1,1,1E-300)', n(-7)],
  ['FLOOR.MATH', 'FLOOR.MATH(-7.1,1,0)', n(-8)],
  ['FLOOR.MATH', 'FLOOR.MATH(7.1,1,0.01)', n(7)],
  ['CEILING.MATH', 'CEILING.MATH(-7.1,1,0.5)', n(-8)],
  ['CEILING.MATH', 'CEILING.MATH(-7.1,1,"0.01")', n(-8)],
  ['CEILING.MATH', 'CEILING.MATH(-7.1,1,TRUE)', n(-8)],
  ['CEILING.MATH', 'CEILING.MATH(-7.1,1,FALSE)', n(-7)],
  ['CEILING.MATH', 'CEILING.MATH(7.1,1,0.01)', n(8)],
  // ROUND just below a half, within the 15th digit and past it
  ['ROUND', 'ROUND(2.5-2^-51,0)', n(3)],
  ['ROUND', 'ROUND(-2.5+2^-51,0)', n(-3)],
  ['ROUND', 'ROUND(1234.5-2^-38,0)', n(1235)],
  ['ROUND', 'ROUND(1234.5-2^-30,0)', n(1234)],
  ['ROUND', 'ROUND(0.125-2^-56,2)', n(0.13)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
