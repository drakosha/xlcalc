/**
 * LET in books with defined names (each case held out: its answer was predicted and committed before Excel
 * computed it): a LET name shadows a defined name of the same word inside its LET, the defined name is read in a
 * value before its LET name is declared and outside the LET, a sheet's own name too; LET's array result spills;
 * the formula text as Excel shows it (the case of a name, an omitted argument).
 */
import { n, one, type SpillCase } from './cases.js'

const NAMES = { sheets: [], names: [{ name: 'nm', ref: "'S'!$A$1" }, { name: 'loc', ref: "'S'!$A$2", scope: 'S' }] }

export function buildSpillCases(): SpillCase[] {
  return [
    one('shadow', 'a LET name over a defined name; the defined name in a value, before its LET name, and outside', { A1: 10, A2: 20, B1: '=LET(nm,5,nm+1)', B2: '=LET(x,nm,x+1)', B3: '=LET(nm,nm+1,nm)', B4: '=LET(nm,5,nm)+nm', B5: '=LET(loc,1,loc)+loc' }, ['A1:B5'], {
      book: NAMES,
      predicted: { 'S!A1': n(10), 'S!A2': n(20), 'S!B1': n(6), 'S!B2': n(11), 'S!B3': n(11), 'S!B4': n(15), 'S!B5': n(21) },
    }),
    one('spill', "LET's arrays spill: a range, a computed array, a range by a name", { A1: 1, A2: 2, A3: 3, C1: '=LET(x,A1:A3,x*2)', D1: '=LET(s,SEQUENCE(1,2),s+10)', D3: '=LET(r,nm,r)' }, ['A1:E4'], {
      book: { sheets: [], names: [{ name: 'nm', ref: "'S'!$A$1:$A$2" }] },
      predicted: { 'S!A1': n(1), 'S!A2': n(2), 'S!A3': n(3), 'S!C1': n(2), 'S!C2': n(4), 'S!C3': n(6), 'S!D1': n(11), 'S!E1': n(12), 'S!D3': n(1), 'S!D4': n(2) },
    }),
    one('text', 'formula text as Excel shows it: names in their own case, omitted arguments', { A1: 1, B1: '=LET(X,1,x+1)', B2: '=IF(A1,,1)', B3: '=LET(x,,x)' }, ['A1:B3'], {
      predicted: { 'S!A1': n(1), 'S!B1': n(2), 'S!B2': n(0), 'S!B3': n(0) },
    }),
  ]
}
