/**
 * FV, PV, XNPV and MIRR as `finance-ironcalc-held`'s misses showed them: (1 + rate)^nper by Excel's power (0^0,
 * 0 to a negative power, a negative base to a fraction, a root of an odd 1/y), XNPV's blank first or later value
 * and date, MIRR of one value and of outflows with zeros. Answers: `test/excel/cases/finance-ironcalc-held-2.json`;
 * `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { e, n } from '../vals.js'

/** Xb: values (A) and dates (B) with blanks: A1 and B4 blank; C, D the same flows with D1 and C4 blank. */
export const SHEETS: SheetCells = {
  Xb: { A2: -100, A3: 50, A4: 60, B1: 45292, B2: 45323, B3: 45352, C1: -100, C2: 50, C3: 60, D2: 45323, D3: 45352, D4: 45383 },
}

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // (1 + rate)^nper by Excel's power
  ['FV', 'FV(-1,0,1,5)', e('NUM')],
  ['PV', 'PV(-1,0,1)', e('NUM')],
  ['FV', 'FV(-1,-0.5,1)', e('DIV0')],
  ['FV', 'FV(-2,0.5,1)', e('NUM')],
  ['FV', 'FV(-9,1/3,1)', n(-0.333333333333333)],
  ['PV', 'PV(-2,1/3,1)', n(1)],
  ['FV', 'FV(-0.5,0,1,5)', n(-5)],
  // XNPV: the first value or date blank, a later one
  ['XNPV', 'XNPV(0.1,Xb!A1:A3,Xb!B1:B3)', e('VALUE')],
  ['XNPV', 'XNPV(0.1,Xb!C1:C3,Xb!D1:D3)', e('VALUE')],
  ['XNPV', 'XNPV(0.1,Xb!A2:A4,Xb!B2:B4)', e('NUM')],
  ['XNPV', 'XNPV(0.1,Xb!C2:C4,Xb!D2:D4)', e('NUM')],
  ['XNPV', 'XNPV(0.1,Xb!A1:A4,Xb!D1:D4)', e('VALUE')],
  // MIRR of one value, outflows with zeros, IRR's guess of -1 with good flows
  ['MIRR', 'MIRR({5},0.1,0.1)', e('DIV0')],
  ['MIRR', 'MIRR({-5},0.1,0.1)', e('DIV0')],
  ['MIRR', 'MIRR({-5,0},0.1,0.1)', n(-1)],
  ['MIRR', 'MIRR({0,-5},0.1,0.1)', n(-1)],
  ['IRR', 'IRR({-100,110},-1)', e('VALUE')],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
