/**
 * Number text at Excel's largest number, as `refs-ironcalc-held`'s miss showed it: just below 1E+308, at it and past
 * it, negative, with sixteen digits, in VALUE, arithmetic and a numeric argument. Answers:
 * `test/excel/cases/refs-ironcalc-held-2.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { e, n } from '../vals.js'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  ['VALUE', '"9.99999999999999E+307"+0', n(9.99999999999999e307)],
  ['VALUE', '"9.999999999999999E+307"+0', n(9.99999999999999e307)],
  ['VALUE', '"1E+308"+0', e('VALUE')],
  ['VALUE', '"-1E+308"+0', e('VALUE')],
  ['VALUE', '"1.7976931348623157E+308"+0', e('VALUE')],
  ['VALUE', 'VALUE("1E+308")', e('VALUE')],
  ['VALUE', 'VALUE("-9.99999999999999E+307")', n(-9.99999999999999e307)],
  ['VALUE', 'VALUE("$1E+308")', e('VALUE')],
  ['ABS', 'ABS("1E+308")', e('VALUE')],
  ['VALUE', 'VALUE("1E+307")*10', n(1e308)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
