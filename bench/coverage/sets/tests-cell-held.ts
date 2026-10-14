/**
 * Held out after `sumx-cell-held` and `sumproduct-held`, which showed for SUMX2MY2 and CORREL that only a reference to
 * one cell is taken first (its error the answer, a blank #VALUE!), argument by argument, and that an array of one goes on
 * to the sizes. The rule predicted for the other functions over two arrays the engine checks so: CHISQ.TEST, T.TEST (before
 * its tails and type), FORECAST, COVARIANCE.S, INTERCEPT and STEYX. Answers: `test/excel/cases/tests-cell-held.json`;
 * `PREDICTED` was committed before Excel ran the set (D!A1 blank, D!B1:B8 numbers).
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { e } from '../vals.js'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  ['CHISQ.TEST', 'CHISQ.TEST({#DIV/0!},{1,2})', e('NA')],
  ['CHISQ.TEST', 'CHISQ.TEST(D!A1,D!B1:B2)', e('VALUE')],
  ['CHISQ.TEST', 'CHISQ.TEST(D!B1:B2,D!A1)', e('VALUE')],
  ['CHISQ.TEST', 'CHISQ.TEST(D!A1,D!B1)', e('VALUE')],
  ['CHISQ.TEST', 'CHISQ.TEST({1,2},{#DIV/0!})', e('NA')],
  ['T.TEST', 'T.TEST(D!A1,D!B1:B3,2,1)', e('VALUE')],
  ['T.TEST', 'T.TEST(D!A1,D!B1:B3,2,2)', e('VALUE')],
  ['T.TEST', 'T.TEST(D!A1,D!B1:B3,3,1)', e('VALUE')],
  ['T.TEST', 'T.TEST({#DIV/0!},D!B1:B3,2,1)', e('NA')],
  ['T.TEST', 'T.TEST({#DIV/0!},D!B1:B3,2,2)', e('DIV0')],
  ['T.TEST', 'T.TEST({#DIV/0!},D!B1:B3,3,2)', e('NUM')],
  ['FORECAST', 'FORECAST(1,D!A1,D!B1:B2)', e('VALUE')],
  ['FORECAST', 'FORECAST(1,{#DIV/0!},{1,2})', e('NA')],
  ['FORECAST.LINEAR', 'FORECAST.LINEAR(1,D!B1:B2,D!A1)', e('VALUE')],
  ['COVARIANCE.S', 'COVARIANCE.S(D!A1,D!B1)', e('VALUE')],
  ['INTERCEPT', 'INTERCEPT({#DIV/0!},{1,2})', e('NA')],
  ['STEYX', 'STEYX(D!A1,D!B1:B3)', e('VALUE')],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
