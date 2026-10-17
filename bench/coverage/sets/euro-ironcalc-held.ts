/**
 * Number text with the euro sign on the cases of IronCalc's tests (`xlsx/tests/calc_tests`, files saved by Excel)
 * that the engine answered otherwise: EXP, SIGN and VALUE of `€100`, `100€` and ` €100 `. Then the rule around them:
 * a space, a sign or parentheses with the euro, groups, an exponent, a percent, the sign twice or with `$`, other
 * currency signs, arithmetic, criteria and a range of SUM. Answers: `test/excel/cases/euro-ironcalc-held.json`;
 * `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { e, n } from '../vals.js'

/** Eu: the euro texts of IronCalc's cells (A1, A2) and numbers for criteria (B1:B3). */
export const SHEETS: SheetCells = {
  Eu: { A1: '€100', A2: ' €100 ', A3: '€5', B1: 100, B2: 200, B3: 100 },
}

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // IronCalc's cases
  ['EXP', 'EXP(Eu!A1)', n(2.68811714181614e43)],
  ['EXP', 'EXP("100€")', n(2.68811714181614e43)],
  ['EXP', 'EXP(Eu!A2)', n(2.68811714181614e43)],
  ['SIGN', 'SIGN(Eu!A1)', n(1)],
  ['SIGN', 'SIGN("100€")', n(1)],
  ['VALUE', 'VALUE("12€")', n(12)],
  // a space, a sign and parentheses, groups, an exponent
  ['VALUE', 'VALUE("€100")', n(100)],
  ['VALUE', 'VALUE("€ 100")', n(100)],
  ['VALUE', 'VALUE("100 €")', n(100)],
  ['VALUE', 'VALUE("-€100")', n(-100)],
  ['VALUE', 'VALUE("€-100")', n(-100)],
  ['VALUE', 'VALUE("-100€")', n(-100)],
  ['VALUE', 'VALUE("(€100)")', n(-100)],
  ['VALUE', 'VALUE("€1,000.50")', n(1000.5)],
  ['VALUE', 'VALUE("1,000.50€")', n(1000.5)],
  ['VALUE', 'VALUE("€1E3")', n(1000)],
  // not numbers: a percent, the sign twice or with `$`, a fraction, the sign alone, other currencies
  ['VALUE', 'VALUE("€100%")', e('VALUE')],
  ['VALUE', 'VALUE("100€%")', e('VALUE')],
  ['VALUE', 'VALUE("€100€")', e('VALUE')],
  ['VALUE', 'VALUE("$100€")', e('VALUE')],
  ['VALUE', 'VALUE("€1 1/2")', e('VALUE')],
  ['VALUE', 'VALUE("€")', e('VALUE')],
  ['VALUE', 'VALUE("£100")', e('VALUE')],
  ['VALUE', 'VALUE("100$")', e('VALUE')],
  // arithmetic, a text argument, a range of SUM, criteria
  ['VALUE', '"€100"+1', n(101)],
  ['SUM', 'SUM(Eu!A3,1)', n(1)],
  ['SUM', 'SUM("€5",1)', n(6)],
  ['COUNTIF', 'COUNTIF(Eu!B1:B3,"€100")', n(2)],
  ['COUNTIF', 'COUNTIF(Eu!B1:B3,">€150")', n(1)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
