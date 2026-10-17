/**
 * COUNTIF and COUNTIFS on the cases of IronCalc's tests (`xlsx/tests/calc_tests/STATISTICAL`, files saved by Excel)
 * that the engine answered otherwise: the criterion `<>` built from an empty cell (`CONCAT("<>",E26)`) over cells
 * holding a formula's empty text, errors, text and blanks, in a range and a whole row. Then the rule around them: `<>`
 * written out and joined to a blank, in SUMIF, AVERAGEIF and COUNTIFS of two ranges, over one cell; and `=` against
 * `""` over the same cells. Answers: `test/excel/cases/countif-ironcalc-held.json`; `PREDICTED` was committed before
 * Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { n } from '../vals.js'

/** Ct: rows 1, 2, 25 and 26 as in IronCalc's COUNTIF.xlsx (L1, G25, H25 and E26 a formula's empty text), and 1 to 14 in row 27. */
export const SHEETS: SheetCells = {
  Ct: {
    A1: 'Range', B1: 23, C1: '23', D1: true, E1: 1, F1: -23, G1: 'hey', H1: false, I1: 123, L1: { formula: '""' },
    B2: 23,
    A25: 'Name', B25: { formula: 'NA()' }, C25: 12, D25: { formula: '1/0' }, E25: '#N/A', F25: 0, G25: { formula: '""' }, H25: { formula: '""' }, J25: 0,
    E26: { formula: '""' },
    ...Object.fromEntries(Array.from({ length: 14 }, (_, i) => [`${String.fromCharCode(66 + i)}27`, i + 1])),
  },
}

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // IronCalc's cases: `<>` and an empty cell, in a range and a whole row
  ['COUNTIF', 'COUNTIF(Ct!$B$1:$O$1, CONCAT("<>",Ct!I2))', n(9)],
  ['COUNTIF', 'COUNTIF(Ct!$1:$1, CONCAT("<>",Ct!I2))', n(10)],
  ['COUNTIF', 'COUNTIF(Ct!$B$25:$H$25, CONCAT("<>", Ct!E26))', n(7)],
  ['COUNTIF', 'COUNTIF(Ct!$25:$25, CONCAT("<>", Ct!E26))', n(9)],
  ['COUNTIFS', 'COUNTIFS(Ct!$B$25:$H$25, CONCAT("<>", Ct!E26))', n(7)],
  ['COUNTIF', 'COUNTIF(Ct!$B$25:$O$25, Ct!E26)', n(8)],
  ['COUNTIF', 'COUNTIF(Ct!$B$1:$O$1, CONCAT("<>",Ct!B2))', n(13)],
  // `<>` written out, joined to a blank, over one cell, in the other functions of criteria
  ['COUNTIF', 'COUNTIF(Ct!B25:O25,"<>")', n(8)],
  ['COUNTIF', 'COUNTIF(Ct!B25:O25,"<>"&"")', n(8)],
  ['COUNTIF', 'COUNTIF(Ct!B25:O25,"<>"&Ct!K1)', n(8)],
  ['COUNTIF', 'COUNTIF(Ct!L1,"<>")', n(1)],
  ['COUNTIF', 'COUNTIF(Ct!K1,"<>")', n(0)],
  ['COUNTIFS', 'COUNTIFS(Ct!B25:O25,"<>",Ct!B1:O1,"<>")', n(7)],
  ['SUMIF', 'SUMIF(Ct!B25:O25,"<>",Ct!B27:O27)', n(37)],
  ['AVERAGEIF', 'AVERAGEIF(Ct!B25:O25,"<>",Ct!B27:O27)', n(4.625)],
  ['SUMIFS', 'SUMIFS(Ct!B27:O27,Ct!B25:O25,"<>",Ct!B1:O1,"<>")', n(28)],
  // `""` takes a formula's empty text for a blank, `=` takes only blanks
  ['COUNTIF', 'COUNTIF(Ct!B25:O25,"")', n(8)],
  ['COUNTIF', 'COUNTIF(Ct!L1,"")', n(1)],
  ['COUNTIF', 'COUNTIF(Ct!B25:O25,"=")', n(6)],
  ['COUNTIF', 'COUNTIF(Ct!L1,"=")', n(0)],
  ['COUNTIFS', 'COUNTIFS(Ct!B25:O25,"=")', n(6)],
  ['SUMIF', 'SUMIF(Ct!B25:O25,"=",Ct!B27:O27)', n(68)],
  ['SUMIF', 'SUMIF(Ct!B25:O25,"",Ct!B27:O27)', n(81)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
