/**
 * Held out after `sumx-cell-held`, whose misses showed that SUMPRODUCT answers an error inside an array before it checks
 * the sizes and a blank cell with #VALUE! even against one number. The rules predicted: SUMX2MY2 and its kin, CORREL and
 * its kin take each argument's cell first, in order (its error the answer, a blank #VALUE!), then the sizes; SUMPRODUCT
 * takes the errors of its arrays first, in order, then a blank cell (#VALUE!), then the sizes (#VALUE!). Answers:
 * `test/excel/cases/sumproduct-held.json`; `PREDICTED` was committed before Excel ran the set (D!A1 blank, D!A7 `1/0`,
 * D!A8 `NA()`, D!A12 0, D!A13 `""`, D!B1:B8 numbers, D!E1:E8 numbers with `1/0` in E3, D!I13:I17 numbers with `NA()`,
 * D!K1:K8 blank).
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { n, e } from '../vals.js'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // a blank cell and an error in a cell: the first argument's cell first
  ['SUMX2MY2', 'SUMX2MY2(D!A1,D!A7)', e('VALUE')],
  ['SUMX2MY2', 'SUMX2MY2(D!A7,D!A1)', e('DIV0')],
  ['SUMX2MY2', 'SUMX2MY2(D!K1:K2,D!A1)', e('VALUE')],
  ['CORREL', 'CORREL(D!A1,D!A8)', e('VALUE')],
  // SUMPRODUCT: errors first, either side and as a third array, inside arrays of other sizes
  ['SUMPRODUCT', 'SUMPRODUCT(D!A1,D!A7)', e('DIV0')],
  ['SUMPRODUCT', 'SUMPRODUCT(D!A7,D!A1)', e('DIV0')],
  ['SUMPRODUCT', 'SUMPRODUCT(D!B1:B2,D!B1:B3,D!A7)', e('DIV0')],
  ['SUMPRODUCT', 'SUMPRODUCT(D!B1:B3,D!E1:E8)', e('DIV0')],
  ['SUMPRODUCT', 'SUMPRODUCT(D!E1:E8,D!I13:I17)', e('DIV0')],
  // SUMPRODUCT: a blank cell alone, against itself, a cell, an array of one, an array; a blank range of two
  ['SUMPRODUCT', 'SUMPRODUCT(D!A1)', e('VALUE')],
  ['SUMPRODUCT', 'SUMPRODUCT(D!A1,D!A1)', e('VALUE')],
  ['SUMPRODUCT', 'SUMPRODUCT(D!A1:A1,D!B1)', e('VALUE')],
  ['SUMPRODUCT', 'SUMPRODUCT(D!A1,{1})', e('VALUE')],
  ['SUMPRODUCT', 'SUMPRODUCT({1,2},D!A1)', e('VALUE')],
  ['SUMPRODUCT', 'SUMPRODUCT(D!K1:K2,D!B1:B2)', n(0)],
  // SUMPRODUCT: a cell of 0 against a range, a cell of empty text against a number
  ['SUMPRODUCT', 'SUMPRODUCT(D!A12,D!B1:B2)', e('VALUE')],
  ['SUMPRODUCT', 'SUMPRODUCT(D!A13,D!B1)', n(0)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)

/** The prediction Excel did not confirm (the Mac and Windows alike): SUMPRODUCT takes its arrays in turn, a blank cell before a later array's error. */
export const MISSED: readonly string[] = ['SUMPRODUCT#1']
