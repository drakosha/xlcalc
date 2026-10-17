/**
 * XMATCH's and XLOOKUP's match mode 3, regular expressions, which Excel for Mac 16.114 has and Excel for Windows
 * 16.0.20430 does not (#VALUE!), after IronCalc's `XMATCH.xlsx` (`XMATCH("[A-Z][a-z]+y",A2:A9,3,1)` is Germany's
 * row): a pattern found in part of a text or the whole, anchors, case and a leading `(?i)`, a pattern that is no
 * regular expression, a number as the key and among the values, from the end, binary search, XLOOKUP's if_not_found.
 * Answers: `test/excel/cases/xmatch-regex-held.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { e, n, s } from '../vals.js'

/** Rx: IronCalc's countries (A1:A8) with numbers beside them, then 42 as a number, `x42` and `GERMANY`. */
export const SHEETS: SheetCells = {
  Rx: {
    A1: 'Albania', A2: 'Indonesia', A3: 'India', A4: 'Germany', A5: 'Spain', A6: 'Venezuela', A7: 'Italy', A8: 'Francialy',
    B1: 1, B2: 2, B3: 3, B4: 4, B5: 5, B6: 6, B7: 7, B8: 8, A10: 42, A11: 'x42', A12: 'GERMANY',
  },
}

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // IronCalc's cases
  ['XMATCH', 'XMATCH("[A-Z][a-z]+y",Rx!A1:A8,3,1)', n(4)],
  ['XMATCH', 'XMATCH("[A-Z][a-z]+y",Rx!A1:A8,3,-1)', n(8)],
  ['INDEX', 'INDEX(Rx!A1:A8,XMATCH("[A-Z][a-z]+y",Rx!A1:A8,3,1))', s('Germany')],
  ['XMATCH', 'XMATCH(42,Rx!A1:A8,3)', e('NA')],
  // part of a text, anchors, from the end
  ['XMATCH', 'XMATCH("an",Rx!A1:A8,3)', n(1)],
  ['XMATCH', 'XMATCH("^In",Rx!A1:A8,3)', n(2)],
  ['XMATCH', 'XMATCH("ia$",Rx!A1:A8,3,-1)', n(3)],
  ['XMATCH', 'XMATCH("y$",Rx!A1:A8,3)', n(4)],
  ['XMATCH', 'XMATCH("I.a",Rx!A1:A8,3)', n(7)],
  ['XMATCH', 'XMATCH("Ind*",Rx!A1:A8,3)', n(2)],
  ['XMATCH', 'XMATCH("",Rx!A1:A8,3)', n(1)],
  // case, a leading (?i), a pattern that is none
  ['XMATCH', 'XMATCH("germany",Rx!A1:A12,3)', e('NA')],
  ['XMATCH', 'XMATCH("(?i)germany",Rx!A1:A12,3)', n(4)],
  ['XMATCH', 'XMATCH("GERMANY",Rx!A1:A12,3)', n(12)],
  ['XMATCH', 'XMATCH("[",Rx!A1:A8,3)', e('VALUE')],
  // numbers among the values and as the key
  ['XMATCH', 'XMATCH(42,Rx!A10:A11,3)', n(2)],
  ['XMATCH', 'XMATCH("\\d",Rx!A10:A12,3)', n(2)],
  // binary search, XLOOKUP
  ['XMATCH', 'XMATCH("Ind.*",Rx!A1:A8,3,2)', e('VALUE')],
  ['XLOOKUP', 'XLOOKUP("^Sp",Rx!A1:A8,Rx!B1:B8,,3)', n(5)],
  ['XLOOKUP', 'XLOOKUP("zzz",Rx!A1:A8,Rx!B1:B8,"none",3)', s('none')],
  ['XLOOKUP', 'XLOOKUP("a$",Rx!A1:A8,Rx!B1:B8,,3,-1)', n(6)],
  ['XLOOKUP', 'XLOOKUP("a",Rx!A1:A8,Rx!B1:B8,,3,-2)', e('VALUE')],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
