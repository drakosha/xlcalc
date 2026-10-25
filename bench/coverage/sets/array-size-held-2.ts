/**
 * What `array-size-held` led to, held out (Excel en-US): how many columns an array may have (SEQUENCE, RANDARRAY,
 * TOROW, HSTACK, WRAPROWS, operators), whether TRANSPOSE and WRAPCOLS may make more rows than a sheet has, and the
 * sheet's last row reached exactly; up to some 2 million values, as there. Answers:
 * `test/excel/cases/array-size-held-2.json`; `PREDICTED`, the engine's answers, was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { n } from '../vals.js'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  ['SEQUENCE', 'COLUMNS(SEQUENCE(1,1048577))', n(1048577)],
  ['SEQUENCE', 'COLUMNS(SEQUENCE(1,2000000))', n(2000000)],
  ['RANDARRAY', 'COLUMNS(RANDARRAY(1,1048577))', n(1048577)],
  ['TOROW', 'COLUMNS(TOROW(SEQUENCE(1100,1000)))', n(1100000)],
  ['HSTACK', 'COLUMNS(HSTACK(SEQUENCE(1,16384),1))', n(16385)],
  ['HSTACK', 'COLUMNS(HSTACK(SEQUENCE(1,1048576),1))', n(1048577)],
  ['TRANSPOSE', 'ROWS(TRANSPOSE(SEQUENCE(1,1048577)))', n(1048577)],
  ['WRAPCOLS', 'ROWS(WRAPCOLS(SEQUENCE(1,1048577),1048577))', n(1048577)],
  ['WRAPROWS', 'COLUMNS(WRAPROWS(SEQUENCE(1,1048577),1048577))', n(1048577)],
  ['SEQUENCE', 'ROWS(SEQUENCE(1048576))', n(1048576)],
  ['VSTACK', 'ROWS(VSTACK(SEQUENCE(1048575),1))', n(1048576)],
  ['TOCOL', 'ROWS(TOCOL(SEQUENCE(1024,1024)))', n(1048576)],
  ['SEQUENCE', 'COLUMNS(SEQUENCE(2,1048577))', n(1048577)],
  ['operators', 'ROWS(SEQUENCE(1048576)*SEQUENCE(1,2))', n(1048576)],
  ['operators', 'COLUMNS(SEQUENCE(1,16385)*SEQUENCE(2))', n(16385)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
