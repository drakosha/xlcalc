/**
 * The database functions on the cases of IronCalc's tests (`xlsx/tests/calc_tests/DATABASE`, files saved by Excel)
 * that the engine answered otherwise: a database whose headers are TRUE, 1 and the text `true`, under criteria headed
 * TRUE, `true` and 1. Then the rule around them: a field named `TRUE`, a criteria header `TRUE` in capitals and the
 * text `1`, a database that starts at the text header, each function of the group. Answers:
 * `test/excel/cases/database-ironcalc-held.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { e, n } from '../vals.js'

/** Db: the sheet of IronCalc's DMIN_DMAX_DAVERAGE_DSUM.xlsx (A1:K11 and the criteria I23:L25), and criteria of the set's own in M. */
export const SHEETS: SheetCells = {
  Db: {
    A1: 'Region', B1: 'Salesperson', C1: 'Quarter', D1: 'Product', E1: 'Units', F1: 'Revenue', G1: 'Date', H1: 'Contractor', I1: true, J1: 1, K1: 'true',
    A2: 'East', B2: 'Alice', C2: 'Q1', D2: 'Widget', E2: 120, F2: 12000, G2: 45306, H2: true, I2: 'a', K2: 'b',
    A3: 'West', B3: 'Bob', C3: 'Q1', D3: 'Gadget', E3: 80, F3: 9600, G3: 45324, H3: false, I3: 'b', J3: 'b', K3: 'b',
    A4: 'East', B4: 'Cara', C4: 'Q2', D4: 'Widget', E4: 200, F4: 20000, G4: 45392, H4: true, I4: 'A', J4: 'B', K4: 'a',
    A5: 'South', B5: 'Dan', C5: 'Q1', D5: 'Widget', E5: 150, F5: 15000, G5: 45369, H5: true, I5: 'A', K5: 'a',
    A6: 'West', B6: 'Ella', C6: 'Q2', D6: 'Gadget', E6: 50, F6: 6000, G6: 45434, H6: false, I6: 'a', J6: 'BB', K6: 12,
    A7: 'East', B7: 'Frank', C7: 'Q3', D7: 'Gizmo', E7: 300, F7: 30000, G7: 45481, H7: true, I7: { formula: 'NA()' }, J7: 'balpha', K7: 1,
    A8: 'North', B8: 'Gina', C8: 'Q3', D8: 'Widget', E8: 100, F8: 11000, G8: 45503, H8: false, I8: 'b', K8: { formula: 'NA()' },
    A9: 'West', B9: 'Hank', C9: 'Q3', D9: 'Widget', E9: 400, F9: 38000, G9: 45515, H9: false, I9: 23, K9: 'b',
    A10: 'East', B10: 'Ivy', C10: 'Q4', D10: 'Gadget', E10: 250, F10: 24000, G10: 45571, H10: true, I10: 455, J10: 23, K10: 'A',
    A11: 'South', B11: 'Jake', C11: 'Q4', D11: 'Widget', E11: 200, F11: 19800, G11: 45621, H11: true, I11: 2, K11: { formula: '1/0' },
    I23: 29, J23: 30, K23: 31, L23: 32,
    I24: true, J24: 'true', K24: 1, L24: 'One',
    I25: 'a', J25: '<>a', K25: 'b', L25: 'aa',
    M24: 'TRUE', M25: 'a', M27: '1', M28: 'b',
  },
}

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // IronCalc's cases: criteria headed TRUE and `true` find the column headed TRUE, the first that reads as either
  ['DSUM', 'DSUM(Db!A1:K11,"Revenue",Db!I24:I25)', n(53000)],
  ['DSUM', 'DSUM(Db!A1:K11,"Revenue",Db!J24:J25)', n(132400)],
  ['DMIN', 'DMIN(Db!A1:K11,"Revenue",Db!I24:I25)', n(6000)],
  ['DMAX', 'DMAX(Db!A1:K11,"Revenue",Db!I24:I25)', n(20000)],
  ['DAVERAGE', 'DAVERAGE(Db!A1:K11,"Revenue",Db!J24:J25)', n(22066.6666666667)],
  ['DVAR', 'DVAR(Db!A1:K11,"Revenue",Db!I24:I25)', n(34250000)],
  // a header 1 and the text `1`, a header `TRUE` in capitals, a field named `TRUE`
  ['DSUM', 'DSUM(Db!A1:K11,"Revenue",Db!K24:K25)', n(65600)],
  ['DSUM', 'DSUM(Db!A1:K11,"Revenue",Db!M27:M28)', n(65600)],
  ['DSUM', 'DSUM(Db!A1:K11,"Revenue",Db!M24:M25)', n(53000)],
  ['DCOUNTA', 'DCOUNTA(Db!A1:K11,"TRUE",Db!K24:K25)', n(4)],
  ['DCOUNT', 'DCOUNT(Db!A1:K11,"Units",Db!J24:J25)', n(6)],
  ['DGET', 'DGET(Db!A1:K11,"Salesperson",Db!I24:I25)', e('NUM')],
  // a database that starts at the text header: TRUE names it; DGET of several records, DSTDEVP, DPRODUCT
  ['DCOUNTA', 'DCOUNTA(Db!J1:K11,2,Db!I24:I25)', n(3)],
  ['DSTDEVP', 'DSTDEVP(Db!A1:K11,"Revenue",Db!J24:J25)', n(10029.7335735081)],
  ['DPRODUCT', 'DPRODUCT(Db!A1:K11,"Units",Db!I24:I25)', n(180000000)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
