/**
 * Match mode 3 over values that are not text, as `xmatch-regex-held`'s misses showed: numbers by their text (a
 * fraction, a large number in scientific notation, a negative one), booleans, a blank and an error among the values.
 * Answers: `test/excel/cases/xmatch-regex-held-2.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { e, n } from '../vals.js'

/** Ry: 1.5, 1E+20, -7, TRUE, a blank (A5), #N/A, `end`. */
export const SHEETS: SheetCells = {
  Ry: { A1: 1.5, A2: 1e20, A3: -7, A4: true, A6: { formula: 'NA()' }, A7: 'end' },
}

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  ['XMATCH', 'XMATCH("1\\.5",Ry!A1:A7,3)', n(1)],
  ['XMATCH', 'XMATCH("E\\+20",Ry!A1:A7,3)', n(2)],
  ['XMATCH', 'XMATCH("^-",Ry!A1:A7,3)', n(3)],
  ['XMATCH', 'XMATCH("TRUE",Ry!A1:A7,3)', n(4)],
  ['XMATCH', 'XMATCH("N/A",Ry!A1:A7,3)', e('NA')],
  ['XMATCH', 'XMATCH("^$",Ry!A1:A7,3)', e('NA')],
  ['XMATCH', 'XMATCH("d$",Ry!A1:A7,3)', n(7)],
  ['XMATCH', 'XMATCH(1.5,Ry!A1:A7,3)', n(1)],
  ['XMATCH', 'XMATCH(TRUE,Ry!A1:A7,3)', n(4)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)

/**
 * The prediction Excel for Mac did not confirm: a blank among the values is empty text (`^$` finds A5). The set
 * `xmatch-regex-held-3` measures blanks.
 */
export const MISSED: readonly string[] = ['XMATCH#6']
