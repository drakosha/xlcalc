/**
 * Held out (Excel en-US): arrays of more rows or columns than a sheet has, which the engine makes in memory as Excel
 * makes MUNIT(45688) (`signatures.ts`), but no larger than some 2 million values, so that Excel's resources hold (an
 * array of 2^31 values, `#VALUE!` here, is not run). Answers: `test/excel/cases/array-size-held.json`; `PREDICTED`, the
 * engine's answers, was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { n } from '../vals.js'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  ['SEQUENCE', 'ROWS(SEQUENCE(1048577))', n(1048577)],
  ['SEQUENCE', 'INDEX(SEQUENCE(1048577),1048577)', n(1048577)],
  ['SEQUENCE', 'COLUMNS(SEQUENCE(1,16385))', n(16385)],
  ['SEQUENCE', 'INDEX(SEQUENCE(1,16385),1,16385)', n(16385)],
  ['SEQUENCE', 'ROWS(SEQUENCE(2000000))', n(2000000)],
  ['SEQUENCE', 'SUM(SEQUENCE(1048577))', n(549757386753)],
  ['SEQUENCE', 'ROWS(SEQUENCE(1048577)*1)', n(1048577)],
  ['SEQUENCE', 'ROWS(SEQUENCE(1048576,2))', n(1048576)],
  ['RANDARRAY', 'ROWS(RANDARRAY(1048577))', n(1048577)],
  ['RANDARRAY', 'COLUMNS(RANDARRAY(1,16385))', n(16385)],
  ['VSTACK', 'ROWS(VSTACK(SEQUENCE(1048576),1))', n(1048577)],
  ['TOCOL', 'ROWS(TOCOL(SEQUENCE(1100,1000)))', n(1100000)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)

/**
 * The predictions Excel did not confirm (Excel for Mac 16.114): SEQUENCE and RANDARRAY of more rows than a sheet has
 * are #VALUE! (16,385 columns are made), and VSTACK and TOCOL past them #NUM!; the engine then followed.
 */
export const MISSED = ['SEQUENCE#1', 'SEQUENCE#2', 'SEQUENCE#5', 'SEQUENCE#6', 'SEQUENCE#7', 'RANDARRAY#1', 'VSTACK#1', 'TOCOL#1']
