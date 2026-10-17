/**
 * TEXT with a digit placeholder in a date's code, on the cases of IronCalc's tests (`xlsx/tests/calc_tests/TEXT.xlsx`,
 * saved by Excel) that the engine answered otherwise: `mm###` for dates and small numbers. Then the rule around it:
 * `0` and `?` after a date code, before it, after hours and elapsed hours, a second's decimals by `0` and by `#`, the
 * placeholder escaped or quoted, a second section of digits. Answers: `test/excel/cases/text-date-ironcalc-held.json`;
 * `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { e, s } from '../vals.js'

/** Tx: a date of IronCalc's sheet (A1 = A9+123 there, 42165) and its format code in B1. */
export const SHEETS: SheetCells = {
  Tx: { A1: { formula: '42042+123' }, B1: 'mm###' },
}

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // IronCalc's cases
  ['TEXT', 'TEXT(41181,"mm###")', e('VALUE')],
  ['TEXT', 'TEXT(1,"mm###")', e('VALUE')],
  ['TEXT', 'TEXT(Tx!A1,Tx!B1)', e('VALUE')],
  // other placeholders after and before a date code, after hours and elapsed hours
  ['TEXT', 'TEXT(41181,"mm0")', e('VALUE')],
  ['TEXT', 'TEXT(41181,"mm?")', e('VALUE')],
  ['TEXT', 'TEXT(41181,"yyyy#")', e('VALUE')],
  ['TEXT', 'TEXT(41181,"0.0 yyyy")', e('VALUE')],
  ['TEXT', 'TEXT(41181,"#,##0 mm")', e('VALUE')],
  ['TEXT', 'TEXT(41181.5,"hh 00")', e('VALUE')],
  ['TEXT', 'TEXT(1.5,"[h] #")', e('VALUE')],
  ['TEXT', 'TEXT(41181.5,"mm.00")', e('VALUE')],
  // a second's decimals, by 0 and by #
  ['TEXT', 'TEXT(41181.5,"hh:mm:ss.00")', s('12:00:00.00')],
  ['TEXT', 'TEXT(41181.5,"ss.000")', s('00.000')],
  ['TEXT', 'TEXT(41181.5,"hh:mm:ss.##")', e('VALUE')],
  ['TEXT', 'TEXT(41181.5,"hh:mm:ss.0#")', e('VALUE')],
  // the placeholder escaped or quoted is a literal; a second section of digits
  ['TEXT', 'TEXT(41181,"mm\\#")', s('09#')],
  ['TEXT', 'TEXT(41181,"mm""#""")', s('09#')],
  ['TEXT', 'TEXT(41181,"mm;###")', s('09')],
  ['TEXT', 'TEXT(-5,"mm;###")', s('5')],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)

/**
 * The prediction Excel did not confirm (the Mac and Windows alike): `mm.00` is minutes and decimals, not a month
 * and a placeholder (`00.00` at noon). The set `text-date-ironcalc-held-2` measures what the decimals are.
 */
export const MISSED: readonly string[] = ['TEXT#11']
