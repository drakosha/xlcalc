/**
 * References, sorting and lifting on the cases of IronCalc's tests (`xlsx/tests/calc_tests`, files saved by Excel)
 * that the engine answered otherwise: ISREF of FILTER's if_empty given as a reference, INDEX of an array's row,
 * SORT's order of errors, IFERROR and LEFT lifted over arrays of different lengths, arithmetic on `1e400`. Then the
 * rules around them: a function reads a position outside a shorter array as #N/A in its turn (IF, IFNA, ROUND, MID),
 * an operator answers #N/A there; INDEX of other rows of an array and of a range; SORT and SORTBY descending.
 * Answers: `test/excel/cases/refs-ironcalc-held.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { b, e, n, s } from '../vals.js'

/** Rf: IronCalc's ISREF.xlsx (L, M), IFERROR.xlsx (G, H) and MID_LEFT_RIGHT_arrays.xlsx (A) cells; errors to sort in B. */
export const SHEETS: SheetCells = {
  Rf: {
    L1: 10, L2: 20, M1: false, M2: false,
    A1: 'Hola', A2: 'Adios', A3: 'Vamos', A4: '  A log text. ', A6: { formula: '1/0' }, A7: { formula: 'NA()' }, A8: { formula: 'SQRT(-1)' }, A9: 'ironcalc',
    B1: { formula: 'NA()' }, B2: { formula: '1/0' }, B3: { formula: 'SQRT(-1)' }, B4: { formula: '"x"+1' }, B5: 5,
    C1: 'c1', C2: 'c2', C3: 'c3', C4: 'c4', C5: 'c5',
    G4: 100, G5: 250, G6: 80, G7: 60, H4: 5, H5: 0, H6: 8, H7: 0, G12: 10, G13: 20,
  },
}

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // IronCalc's cases
  ['ISREF', 'ISREF(FILTER(Rf!L1:L2, Rf!M1:M2, Rf!L1:L2))', b(false)],
  ['FILTER', 'FILTER(Rf!L1:L2, Rf!M1:M2, Rf!L1)+1', n(11)],
  ['INDEX', 'SUM(INDEX({1,2,3;4,5,6},1))', n(6)],
  ['SORT', 'INDEX(SORT(Rf!B1:B5),2)', e('DIV0')],
  ['SORT', 'INDEX(SORT(Rf!B1:B5),5)', e('NA')],
  ['IFERROR', 'INDEX(IFERROR(Rf!G4:G7/Rf!H4:H7, Rf!G12:G13),3)', n(10)],
  ['IFERROR', 'INDEX(IFERROR(Rf!G4:G7/Rf!H4:H7, Rf!G12:G13),4)', e('NA')],
  ['LEFT', 'INDEX(LEFT(Rf!A1:A9,SEQUENCE(5)),6)', e('DIV0')],
  ['LEFT', 'INDEX(LEFT(Rf!A1:A9,SEQUENCE(5)),8)', e('NUM')],
  ['LEFT', 'INDEX(LEFT(Rf!A1:A9,SEQUENCE(5)),9)', e('NA')],
  ['VALUE', '"1e400"+1', e('VALUE')],
  // FILTER's if_empty of one cell, a value too
  ['ISREF', 'ISREF(FILTER(Rf!L1:L2,Rf!M1:M2,Rf!L1))', b(false)],
  // INDEX of an array's rows, of a range
  ['INDEX', 'SUM(INDEX({1,2,3;4,5,6},2))', n(15)],
  ['INDEX', 'COLUMNS(INDEX({1,2,3;4,5,6},1))', n(3)],
  ['INDEX', 'INDEX({1,2,3;4,5,6},3)', e('REF')],
  ['INDEX', 'SUM(INDEX({1,2,3;4,5,6},0))', n(21)],
  ['INDEX', 'INDEX({1,2,3;4,5,6},-1)', e('VALUE')],
  ['INDEX', 'INDEX(Rf!G4:H7,1)', e('REF')],
  // errors sorted by ERROR.TYPE's number, both ways, as keys of SORTBY
  ['SORT', 'INDEX(SORT(Rf!B1:B5),3)', e('VALUE')],
  ['SORT', 'INDEX(SORT(Rf!B1:B5),4)', e('NUM')],
  ['SORT', 'INDEX(SORT(Rf!B1:B5,1,-1),1)', e('NA')],
  ['SORT', 'INDEX(SORT(Rf!B1:B5,1,-1),4)', e('DIV0')],
  ['SORT', 'INDEX(SORT(Rf!B1:B5,1,-1),5)', n(5)],
  ['SORTBY', 'INDEX(SORTBY(Rf!C1:C5,Rf!B1:B5),2)', s('c2')],
  ['SORTBY', 'INDEX(SORTBY(Rf!C1:C5,Rf!B1:B5,-1),2)', s('c3')],
  // a function reads a position outside a shorter array as #N/A in its turn; an operator answers #N/A there
  ['IF', 'INDEX(IF({TRUE;FALSE;FALSE},{1;2},3),3)', n(3)],
  ['IFNA', 'INDEX(IFNA({1;2;3},{7;8}),3)', n(3)],
  ['IFERROR', 'INDEX(IFERROR({1;2;3},{7;8}),3)', n(3)],
  ['ROUND', 'INDEX(ROUND({1.25;2.25;3.25},{1;1}),3)', e('NA')],
  ['MID', 'INDEX(MID({"abc";"def";"ghi"},{1;2},1),3)', e('NA')],
  ['LEFT', 'INDEX(LEFT(Rf!A6:A8,{1;2}),3)', e('NUM')],
  ['LEFT', 'INDEX(Rf!A6:A8&{"a";"b"},3)', e('NA')],
  // number text past the largest number
  ['VALUE', 'VALUE("1e400")', e('VALUE')],
  ['VALUE', '"1e308"+0', n(1e308)],
  ['ISNUMBER', 'ISNUMBER(--"1e400")', b(false)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
