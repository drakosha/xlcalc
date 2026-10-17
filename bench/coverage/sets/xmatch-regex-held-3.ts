/**
 * Match mode 3 over blanks, as `xmatch-regex-held-2`'s miss showed: a blank inside the values and past their
 * occupied part (a whole column, from the start and from the end), a pattern that matches empty text and one that
 * does not. Answers: `test/excel/cases/xmatch-regex-held-3.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { e, n } from '../vals.js'

/** Rz: `a`, a blank (A2), `b`, then blanks down the column. */
export const SHEETS: SheetCells = { Rz: { A1: 'a', A3: 'b' } }

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  ['XMATCH', 'XMATCH("^$",Rz!A1:A3,3)', n(2)],
  ['XMATCH', 'XMATCH("^$",Rz!A1:A3,3,-1)', n(2)],
  ['XMATCH', 'XMATCH("^$",Rz!A:A,3,-1)', n(1048576)],
  ['XMATCH', 'XMATCH("^$",Rz!A3:A10,3)', n(2)],
  ['XMATCH', 'XMATCH("x*",Rz!A1:A3,3)', n(1)],
  ['XMATCH', 'XMATCH(".",Rz!A2:A2,3)', e('NA')],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
