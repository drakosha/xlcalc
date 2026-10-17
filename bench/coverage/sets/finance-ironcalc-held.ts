/**
 * FV, PV, IRR, MIRR and XNPV on the cases of IronCalc's tests (`xlsx/tests/calc_tests`, files saved by Excel) that
 * the engine answered otherwise: FV at a rate of -1 over a negative number of periods, IRR with a guess of -1 or
 * less, MIRR of outflows alone, XNPV with a blank among its values. Then the rules around them: other rates and
 * periods, PV likewise, guesses a hair above -1 and with flows of one sign or an error, MIRR of inflows alone and of
 * zeros, XNPV with its first value blank, all blank, a blank date. Answers:
 * `test/excel/cases/finance-ironcalc-held.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { e, n } from '../vals.js'

/**
 * Fi: rows of IronCalc's IRR.xlsx (1, 2 from rows 3 and 5), MIRR.xlsx (3 to 5) and XIRR.xlsx (6, 7 from rows 7 and 20,
 * dates in H:M); Xv: values with blanks (A, C), dates (B), values and dates with a blank date (D, E).
 */
export const SHEETS: SheetCells = {
  Fi: {
    A1: -1000, B1: 230, C1: 200, D1: 125, E1: 200, F1: 10,
    A2: -20, B2: 5, C2: 5, E2: -10, F2: 5,
    A3: -1, B3: -1, C3: -1, D3: -1, E3: -1, F3: -1,
    A4: -6, B4: -1, C4: -2,
    A5: -5, B5: -3,
    A6: 20, B6: 5, C6: 5, E6: 5, F6: 5, H6: 44958, I6: 44986, J6: 45017, K6: 45078, L6: 45108, M6: 45139,
    A7: -20, B7: 5, C7: 5, F7: 5, H7: 45292, I7: 45323, J7: 45353, K7: 45386, L7: 45782, M7: 46179,
  },
  Xv: { A1: -100, A3: 50, A4: 60, B1: 45292, B2: 45323, B3: 45352, B4: 45383, D1: -100, D2: 50, D3: 60, E1: 45292, E3: 45352 },
}

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // IronCalc's cases
  ['FV', 'FV(-1,-2,1)', e('DIV0')],
  ['FV', 'FV(-1,-1,1)', e('DIV0')],
  ['IRR', 'IRR(Fi!A1:F1,-10)', e('VALUE')],
  ['IRR', 'IRR(Fi!A2:F2,-1)', e('VALUE')],
  ['MIRR', 'MIRR(Fi!A3:F3,0.1,0.2)', n(-1)],
  ['MIRR', 'MIRR(Fi!A3:E3,0.1,2)', n(-1)],
  ['MIRR', 'MIRR(Fi!A4:C4,0.1,0.2)', n(-1)],
  ['MIRR', '-MIRR(Fi!A5:B5,0.1,0.2)', n(1)],
  ['XNPV', 'XNPV(0.1,Fi!A6:F6,Fi!H6:M6)', e('NUM')],
  ['XNPV', 'XNPV(1,Fi!A7:F7,Fi!H7:M7)', e('NUM')],
  // FV and PV at a rate of -1 and below
  ['FV', 'FV(-1,2,1)', n(-1)],
  ['FV', 'FV(-1,0,1)', n(0)],
  ['FV', 'FV(-1,-2,1,0,1)', e('DIV0')],
  ['FV', 'FV(-2,-2,1)', n(0)],
  ['FV', 'FV(-2,-1.5,1)', e('NUM')],
  ['PV', 'PV(-1,-2,1)', e('DIV0')],
  ['PV', 'PV(-1,2,1)', e('DIV0')],
  ['PV', 'PV(-1,0,1,5)', n(-5)],
  // IRR's guess against the flows' errors and signs
  ['IRR', 'IRR({-100,110},-0.99)', n(0.1)],
  ['IRR', 'IRR({-100,110},-1.5)', e('VALUE')],
  ['IRR', 'IRR({100,110},-2)', e('VALUE')],
  ['IRR', 'IRR({100,110},0.1)', e('NUM')],
  ['IRR', 'IRR({-100,#N/A},-2)', e('VALUE')],
  // MIRR of inflows alone, of zeros, of outflows and zeros
  ['MIRR', 'MIRR({1,2,3},0.1,0.1)', e('DIV0')],
  ['MIRR', 'MIRR({0,0,0},0.1,0.1)', e('DIV0')],
  ['MIRR', 'MIRR({-1,0,0},0.1,0.1)', n(-1)],
  ['MIRR', 'MIRR({-1,-2,-3},0.1,0.1)', n(-1)],
  // XNPV with blanks: a value between, the first value, all of them, a date
  ['XNPV', 'XNPV(0.1,Xv!A1:A4,Xv!B1:B4)', e('NUM')],
  ['XNPV', 'XNPV(0.1,Xv!A2:A4,Xv!B2:B4)', e('NUM')],
  ['XNPV', 'XNPV(0.1,Xv!C1:C2,Xv!B1:B2)', e('VALUE')],
  ['XNPV', 'XNPV(0.1,Xv!D1:D3,Xv!E1:E3)', e('VALUE')],
  ['XNPV', 'XNPV(0.1,Xv!D1:D3,Xv!B1:B3)', n(8.66417201947206)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)

/**
 * The predictions Excel did not confirm (the Mac and Windows alike): at a rate of -1, FV and PV over 0 periods are
 * #NUM!, (1 + rate)^nper taken as Excel's power takes 0^0; XNPV with its first value blank is #VALUE! and with a
 * blank date after the first #NUM!, so the first value or date decides. The set `finance-ironcalc-held-2` measures
 * the rule they led to.
 */
export const MISSED: readonly string[] = ['FV#4', 'PV#3', 'XNPV#4', 'XNPV#6']
