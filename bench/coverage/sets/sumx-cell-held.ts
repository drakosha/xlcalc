/**
 * Held out after `binom-sumx-held`, whose misses showed that an array of one error is no cell (`SUMX2MY2({#DIV/0!},{1,2})`
 * is #N/A, the sizes first), that a blank cell against a range is #VALUE! (`SUMX2MY2(D!A1,D!B1:B2)`) and that SUMPRODUCT
 * answers the error of one cell before its sizes. The rule predicted: a reference to one cell is taken first, its error
 * the answer and a blank #VALUE! (SUMPRODUCT counting a blank as 0 when the sizes agree); anything else, an array of one
 * among them, goes on to the sizes. CORREL and its kin are predicted to take arrays of one and blank cells the same way.
 * Answers: `test/excel/cases/sumx-cell-held.json`; `PREDICTED` was committed before Excel ran the set (D!A1 blank, D!A2
 * the text 2, D!A4 TRUE, D!A7 `1/0`, D!A8 `NA()`, D!A12 0, D!A13 `""`, D!B1:B8 numbers, D!E1:E8 numbers with `1/0` in E3,
 * D!K1:K8 blank).
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { n, e } from '../vals.js'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // one cell of each kind against a range of two: a number, a boolean, empty text and text go on to the sizes
  ['SUMX2MY2', 'SUMX2MY2(D!A12,D!B1:B2)', e('NA')],
  ['SUMX2MY2', 'SUMX2MY2(D!A4,D!B1:B2)', e('NA')],
  ['SUMX2MY2', 'SUMX2MY2(D!A13,D!B1:B2)', e('NA')],
  ['SUMX2MY2', 'SUMX2MY2(D!A2,D!B1:B2)', e('NA')],
  // a blank cell, either side, each function, against a range, a cell and itself; a blank range of two
  ['SUMX2MY2', 'SUMX2MY2(D!B1:B2,D!A1)', e('VALUE')],
  ['SUMXMY2', 'SUMXMY2(D!A1,D!B1:B2)', e('VALUE')],
  ['SUMX2PY2', 'SUMX2PY2(D!A1,{1,2})', e('VALUE')],
  ['SUMX2MY2', 'SUMX2MY2(D!K1,D!B1:B2)', e('VALUE')],
  ['SUMX2MY2', 'SUMX2MY2(D!A1:A1,D!B1:B2)', e('VALUE')],
  ['SUMX2MY2', 'SUMX2MY2(D!A1,D!A1)', e('VALUE')],
  ['SUMX2MY2', 'SUMX2MY2(D!A1,D!B1)', e('VALUE')],
  ['SUMX2MY2', 'SUMX2MY2(D!B1,D!A1)', e('VALUE')],
  ['SUMX2MY2', 'SUMX2MY2(D!K1:K2,D!B1:B2)', e('DIV0')],
  // a cell INDEX gives; an array of one error against a cell's error and against an array of two; values given directly
  ['SUMX2MY2', 'SUMX2MY2(INDEX(D!A7:A8,1),{1,2})', e('DIV0')],
  ['SUMX2MY2', 'SUMX2MY2({#N/A},D!A7)', e('DIV0')],
  ['SUMX2MY2', 'SUMX2MY2({1,2},{#DIV/0!})', e('NA')],
  ['SUMX2MY2', 'SUMX2MY2(5,{1,2})', e('NA')],
  ['SUMX2MY2', 'SUMX2MY2("a",{1,2})', e('NA')],
  ['SUMX2MY2', 'SUMX2MY2(TRUE,D!B1:B2)', e('NA')],
  // CORREL and its kin: an array of one error, a blank cell
  ['CORREL', 'CORREL({#DIV/0!},{1,2})', e('NA')],
  ['CORREL', 'CORREL(D!A1,D!B1:B2)', e('VALUE')],
  ['CORREL', 'CORREL(D!A1,D!B1)', e('VALUE')],
  ['PEARSON', 'PEARSON(D!B1:B2,D!A1)', e('VALUE')],
  ['RSQ', 'RSQ({#DIV/0!},D!B1:B2)', e('NA')],
  ['SLOPE', 'SLOPE(D!A1,D!B1:B2)', e('VALUE')],
  ['COVARIANCE.P', 'COVARIANCE.P(D!A7,D!B1:B2)', e('DIV0')],
  // SUMPRODUCT: a cell's error before the sizes, an array of one after them
  ['SUMPRODUCT', 'SUMPRODUCT(D!B1:B2,D!A7)', e('DIV0')],
  ['SUMPRODUCT', 'SUMPRODUCT(D!A8,{1,2,3})', e('NA')],
  ['SUMPRODUCT', 'SUMPRODUCT(D!A8:A8,D!B1:B3)', e('NA')],
  ['SUMPRODUCT', 'SUMPRODUCT(D!A7,D!A8)', e('DIV0')],
  ['SUMPRODUCT', 'SUMPRODUCT({#DIV/0!},{1,2})', e('VALUE')],
  ['SUMPRODUCT', 'SUMPRODUCT(D!E1:E8,D!B1:B2)', e('VALUE')],
  ['SUMPRODUCT', 'SUMPRODUCT(D!A1,D!B1:B2)', e('VALUE')],
  ['SUMPRODUCT', 'SUMPRODUCT(D!A1,D!B1)', n(0)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)

/**
 * The predictions Excel did not confirm (the Mac and Windows alike): SUMPRODUCT answers an error inside an array, an
 * array of one too, before it compares the sizes, and a blank cell with #VALUE! even against one number. The set
 * `sumproduct-held` measured the order these left open.
 */
export const MISSED: readonly string[] = ['SUMPRODUCT#5', 'SUMPRODUCT#6', 'SUMPRODUCT#8']
