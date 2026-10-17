/**
 * Number text with the euro sign as `euro-ironcalc-held`'s misses showed it: fractions before and after the sign,
 * with a sign and a space, against `$` and parentheses; criteria that take `€100` for text, against cells holding
 * it, the number and its plain text. Answers: `test/excel/cases/euro-ironcalc-held-2.json`; `PREDICTED` was
 * committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { e, n } from '../vals.js'

/** Ec: `€100` as text, 100, `100` as text; numbers beside them for SUMIF. */
export const SHEETS: SheetCells = {
  Ec: { A1: '€100', A2: 100, A3: '100', B1: 1, B2: 2, B3: 4 },
}

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // fractions with the euro sign
  ['VALUE', 'VALUE("1 1/2€")', n(1.5)],
  ['VALUE', 'VALUE("-€1 1/2")', n(-1.5)],
  ['VALUE', 'VALUE("€ 2 3/4")', n(2.75)],
  ['VALUE', 'VALUE("€1 1/2")+0', n(1.5)],
  ['VALUE', 'VALUE("$1 1/2")', e('VALUE')],
  ['VALUE', 'VALUE("(1 1/2€)")', e('VALUE')],
  ['VALUE', 'VALUE("1e3€")', n(1000)],
  // criteria: `€100` is text, a number criterion takes no `€100` in a cell
  ['COUNTIF', 'COUNTIF(Ec!A1:A3,100)', n(2)],
  ['COUNTIF', 'COUNTIF(Ec!A1:A3,"€100")', n(1)],
  ['COUNTIF', 'COUNTIF(Ec!A1:A3,"<>€100")', n(2)],
  ['SUMIF', 'SUMIF(Ec!A1:A3,"100",Ec!B1:B3)', n(6)],
  ['COUNTIFS', 'COUNTIFS(Ec!A1:A3,"€100",Ec!B1:B3,1)', n(1)],
  ['MATCH', 'MATCH("€100",Ec!A1:A3,0)', n(1)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)

/**
 * The predictions Excel did not confirm (the Mac and Windows alike): a fraction is read after `$` and inside
 * parentheses too, and a number criterion's equality takes a cell's `€100` for 100, though the criterion's own text
 * `€100` is text. The set `euro-ironcalc-held-3` measures the rule they led to.
 */
export const MISSED: readonly string[] = ['VALUE#5', 'VALUE#6', 'COUNTIF#1', 'SUMIF#1']
