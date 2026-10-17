/**
 * Number text as `euro-ironcalc-held-2`'s misses showed it: fractions after `$`, inside parentheses, with a sign and
 * a percent; cells holding the euro sign, `$`, parentheses and a fraction under number criteria, against the
 * criterion's own text. Answers: `test/excel/cases/euro-ironcalc-held-3.json`; `PREDICTED` was committed before Excel
 * ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { e, n } from '../vals.js'

/** Ed: texts that read as numbers, 100 three ways, -100 and 1.5; numbers beside them for SUMIF. */
export const SHEETS: SheetCells = {
  Ed: { A1: '100€', A2: ' €100 ', A3: '$100', A4: '(€100)', A5: '€1 1/2', B1: 1, B2: 2, B3: 4, B4: 8, B5: 16 },
}

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // fractions after `$`, in parentheses, with a sign and a percent
  ['VALUE', 'VALUE("$ 1 1/2")', n(1.5)],
  ['VALUE', 'VALUE("-$1 1/2")', n(-1.5)],
  ['VALUE', 'VALUE("(1 1/2)")', n(-1.5)],
  ['VALUE', 'VALUE("($1 1/2)")', n(-1.5)],
  ['VALUE', 'VALUE("1 1/2%")', n(0.015)],
  ['VALUE', 'VALUE("$1 1/2%")', e('VALUE')],
  ['VALUE', '"1 1/2"+"$1 1/2"', n(3)],
  ['ABS', 'ABS("(2 1/4€)")', n(2.25)],
  // number criteria over cells holding the euro sign, `$`, parentheses and a fraction
  ['COUNTIF', 'COUNTIF(Ed!A1:A5,100)', n(3)],
  ['COUNTIF', 'COUNTIF(Ed!A1:A5,-100)', n(1)],
  ['COUNTIF', 'COUNTIF(Ed!A1:A5,1.5)', n(1)],
  ['COUNTIF', 'COUNTIF(Ed!A1:A5,"=100")', n(3)],
  ['COUNTIF', 'COUNTIF(Ed!A1:A5,"=€100")', n(0)],
  ['COUNTIF', 'COUNTIF(Ed!A1:A5,">50")', n(0)],
  ['SUMIF', 'SUMIF(Ed!A1:A5,100,Ed!B1:B5)', n(7)],
  ['COUNTIFS', 'COUNTIFS(Ed!A1:A5,"$100")', n(3)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
