/**
 * The reference operators between references: the intersection, a space (`D!B1:C3 D!C2:D4` is C2:C3, `#NULL!` where
 * the ranges do not meet), and the range operator `:` between references that are not one range token
 * (`D!B1:INDEX(D!B1:B8,3)`): precedence, spaces that are no operator (around operators, after a parenthesis, before
 * a comma), unions, references from INDEX, references across sheets. Answers: `test/excel/cases/intersect.json`;
 * `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { SHEETS as AREA_SHEETS } from './areas.js'
import { n, s, b, e } from '../vals.js'

/** Jan, Feb and Mar of the set `areas`. */
export const SHEETS: SheetCells = AREA_SHEETS

/** Group, formula and the answer predicted before Excel ran (undefined: not predicted). D!B1:B3 4, 8, 15; D!C1:C3 2, 7, 11. */
const CASES: [string, string, Val?][] = [
  ['space', 'SUM(D!B1:C3 D!C2:D4)', n(18)],
  ['space', 'D!B1:C3 D!C2:C2', n(7)],
  ['space', 'D!B1:B3 D!C1:C3', e('NULL')],
  ['space', 'ERROR.TYPE(D!B1:B3 D!C1:C3)', n(1)],
  ['space', 'ROWS(D!B1:C3 D!C2:D4)', n(2)],
  ['space', 'COLUMNS(D!B1:C3 D!B2:D2)', n(2)],
  ['space', 'D!B:B D!3:3', n(15)],
  ['space', 'SUM(D!B1:C3  D!C2:D4)', n(18)],
  ['space', 'SUM((D!B1:C3) (D!C2:D4))', n(18)],
  ['space', 'SUM(D!B1:C3 (D!C2:D4))', n(18)],
  ['space', '-D!B1:C3 D!C2:C2', n(-7)],
  ['space', 'D!B1:C3 D!C2:C2*2', n(14)],
  ['space', 'SUM(D!B1:C3 D!C2:D4+1)', n(20)],
  ['space', 'SUM((D!B1:B3 D!B2:B4,D!C1))', n(25)],
  ['space', 'SUM((D!B1:B3,D!C1:C3) D!B2:C2)', n(15)],
  ['space', 'INDEX(D!B1:C3,0,1) D!B2:C2', n(8)],
  ['space', 'SUM(D!B1:C3 INDEX(D!B1:D4,0,2))', n(20)],
  ['space', 'D!B1:B3 Jan!A1:A3', e('VALUE')],
  ['space', 'ISREF(D!B1:C3 D!C2:D4)', b(true)],
  ['space', 'ISREF(D!B1:B3 D!C1:C3)', b(false)],
  ['space', 'SUM(D!B1:C3 D!C2:D4 D!C3:E5)', n(11)],
  ['space', 'COUNT(D!B1:B3 D!C1:C3)', n(0)],
  ['space', 'IFERROR(D!B1:B3 D!C1:C3,"none")', s('none')],
  ['space', 'SUM(D!B2:B3 D!B2:B3)', n(23)],
  ['space', 'D!A:A D!1:1', n(0)],
  ['space', 'MATCH(11,D!B1:C3 D!C1:C3,0)', n(3)],
  ['space', '@D!B1:C3 D!C2:C2'],
  ['space', 'SUM(Jan:Mar!A1:A3 Jan:Mar!A2:B2)'],
  ['space', 'SUM(D!B1:C3\nD!C2:D4)'],
  // spaces that are no operator
  ['blank', 'D!B1 + D!B2', n(12)],
  ['blank', 'SUM( D!B1:B3 )', n(27)],
  ['blank', 'SUM(D!B1:B3 ,D!C1)', n(29)],
  ['blank', 'SUM(D!B1:B3, D!C1)', n(29)],
  ['blank', '( D!B1 )*2', n(8)],
  // the literal and its error
  ['null', '#NULL!', e('NULL')],
  ['null', 'ISERR(#NULL!)', b(true)],
  ['null', 'IF(ISERROR(D!B1:B3 D!C1:C3),ERROR.TYPE(#NULL!),0)', n(1)],
  // the range operator between references
  ['range', 'SUM(D!B1:INDEX(D!B1:B8,3))', n(27)],
  ['range', 'SUM(INDEX(D!B1:B8,2):D!B4)', n(39)],
  ['range', 'ROWS(D!B1:INDEX(D!B1:B8,5))', n(5)],
  ['range', 'SUM(D!B1:B2:C3)', n(47)],
  ['range', 'SUM(D!B1:D!C3)', n(47)],
  ['range', 'SUM(D!B1:Jan!A1)', e('VALUE')],
  ['range', 'SUM(D!B1:C3 D!C2:INDEX(D!C1:C8,4))', n(18)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)

/** The predictions Excel did not confirm: C3 without a sheet is on the formula's sheet, so `D!B1:B2:C3` joins two sheets. */
export const MISSED: readonly string[] = ['range#4']
