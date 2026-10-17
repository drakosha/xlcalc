/**
 * TEXTAFTER and TEXTBEFORE with delimiters all longer than the text, with and without an instance, as a recorded
 * case (`TEXTBEFORE("A","aBc")` is #N/A) and `textafter-ironcalc-held` part: the plain call, the instance left out
 * between commas with if_not_found or the mode given, the instance given as 1, -1 and from a cell, several
 * delimiters. Answers: `test/excel/cases/textafter-ironcalc-held-2.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { e, s } from '../vals.js'

/** Ta: 1 for an instance read from a cell, a blank in A2. */
export const SHEETS: SheetCells = { Ta: { A1: 1 } }

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  ['TEXTAFTER', 'TEXTAFTER("abc","abcd")', e('NA')],
  ['TEXTBEFORE', 'TEXTBEFORE("abc","abcd")', e('NA')],
  ['TEXTAFTER', 'TEXTAFTER("abc",{"abcd","wxyz"})', e('NA')],
  ['TEXTAFTER', 'TEXTAFTER("abc","abcd",,,,"-")', s('-')],
  ['TEXTAFTER', 'TEXTAFTER("abc","ABCD",,1)', e('NA')],
  ['TEXTAFTER', 'TEXTAFTER("abc","abcd",1)', e('VALUE')],
  ['TEXTBEFORE', 'TEXTBEFORE("abc","abcd",-1)', e('VALUE')],
  ['TEXTAFTER', 'TEXTAFTER("abc","abcd",Ta!A1)', e('VALUE')],
  ['TEXTAFTER', 'TEXTAFTER("abc","abcd",Ta!A2)', e('VALUE')],
  ['TEXTAFTER', 'TEXTAFTER("abc","abcd",1,,,"-")', e('VALUE')],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
