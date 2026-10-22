/**
 * What `values-held-2`'s miss led to, held out (Excel en-US): DOLLAR(1E+190,0), 255 characters, was #VALUE!, so the
 * positive amount's padding of `_)` in the currency format is taken to count; a negative amount in parentheses at 255
 * characters and a positive one at 254 tell that rule from a cap of 254. Answers: `test/excel/cases/values-held-3.json`;
 * `PREDICTED`, the engine's answers, was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { n, e } from '../vals.js'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  ['DOLLAR', 'LEN(DOLLAR(1E+189,0))', n(254)],
  ['DOLLAR', 'LEN(DOLLAR(1E+188,1))', n(254)],
  ['DOLLAR', 'DOLLAR(1E+187,3)', e('VALUE')],
  ['DOLLAR', 'LEN(DOLLAR(-1E+187,1))', n(255)],
  ['DOLLAR', 'LEN(DOLLAR(-1E+186,2))', n(255)],
  ['DOLLAR', 'DOLLAR(-1E+186,3)', e('VALUE')],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
