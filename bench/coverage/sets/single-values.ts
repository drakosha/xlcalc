/**
 * One value where a function takes a table, an array or a list of dates, around the differences the set `omit`
 * found: text as the table of VLOOKUP, HLOOKUP and MATCH; one value as XLOOKUP's lookup array with a longer return
 * array; WRAPROWS and WRAPCOLS of fewer values than the count; text as the arrays of F.TEST and MIRR; text and
 * booleans as XIRR's and XNPV's values and dates. Answers: `test/excel/cases/single-values.json`; `PREDICTED` was
 * committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { Probe } from '../probes.js'
import { collector } from './arrays/common.js'

const n = (v: number): Val => ({ t: 'n', v })
const s = (v: string): Val => ({ t: 's', v })
const e = (v: Extract<Val, { t: 'e' }>['v']): Val => ({ t: 'e', v })

/** Function, formula and the answer predicted before Excel ran (undefined: not predicted). */
const CASES: [string, string, Val?][] = [
  // text as the table: #VALUE!, whatever the text; a number or a boolean is a table of one value
  ['MATCH', 'MATCH(5,"x",0)', e('VALUE')],
  ['MATCH', 'MATCH(5,"5",0)', e('VALUE')],
  ['MATCH', 'MATCH("x","x",0)', e('VALUE')],
  ['MATCH', 'MATCH(5,"",1)', e('VALUE')],
  ['MATCH', 'MATCH(5,""&"",0)', e('VALUE')],
  ['MATCH', 'MATCH(5,5,0)', n(1)],
  ['MATCH', 'MATCH(TRUE,TRUE,0)', n(1)],
  ['MATCH', 'MATCH(5,{"a"},0)', e('NA')],
  ['MATCH', 'MATCH(5,D!A3,0)', e('NA')],
  ['VLOOKUP', 'VLOOKUP("x","x",1,FALSE)', e('VALUE')],
  ['VLOOKUP', 'VLOOKUP(5,"",1,TRUE)', e('VALUE')],
  ['VLOOKUP', 'VLOOKUP(5,5,1,FALSE)', n(5)],
  ['HLOOKUP', 'HLOOKUP(5,"5",1,FALSE)', e('VALUE')],
  ['XMATCH', 'XMATCH(5,"x")', e('NA')],
  ['LOOKUP', 'LOOKUP(5,"x")', e('NA')],
  // one value as XLOOKUP's lookup array: a row of one column, so a return array of one column gives its column
  ['XLOOKUP', 'ARRAYTOTEXT(XLOOKUP(0,0,D!M1:M6),1)', s('{"one";"two";"three";"five";"eight";"thirteen"}')],
  ['XLOOKUP', 'ARRAYTOTEXT(XLOOKUP(1,D!L1,D!M1:M6),1)', s('{"one";"two";"three";"five";"eight";"thirteen"}')],
  ['XLOOKUP', 'XLOOKUP(3,D!L1,D!M1:M6)', e('NA')],
  ['XLOOKUP', 'XLOOKUP(3,0,D!M1:M6,"none")', s('none')],
  ['XLOOKUP', 'XLOOKUP(3,0,D!M1:N6)'],
  ['XLOOKUP', 'XLOOKUP(0,0,D!M1:N6)'],
  ['XLOOKUP', 'ARRAYTOTEXT(XLOOKUP(1,D!L1,D!M1:N1),1)', s('{"one",10}')],
  ['XLOOKUP', 'XLOOKUP(3,{1,2},D!M1:M6)', e('VALUE')],
  ['XLOOKUP', 'XLOOKUP(3,{1;2},D!M1:M6)', e('VALUE')],
  ['XMATCH', 'XMATCH(0,0)', n(1)],
  // fewer values than the count: one line of the values, not padded
  ['WRAPROWS', 'ARRAYTOTEXT(WRAPROWS({1,2,3},5,0),1)', s('{1,2,3}')],
  ['WRAPROWS', 'ARRAYTOTEXT(WRAPROWS({1,2,3},5),1)', s('{1,2,3}')],
  ['WRAPROWS', 'ARRAYTOTEXT(WRAPROWS({7},3),1)', s('{7}')],
  ['WRAPROWS', 'ARRAYTOTEXT(WRAPROWS(D!B1:B2,3),1)', s('{4,8}')],
  ['WRAPROWS', 'ARRAYTOTEXT(WRAPROWS({1,2,3},3),1)', s('{1,2,3}')],
  ['WRAPROWS', 'ARRAYTOTEXT(WRAPROWS({1,2,3,4},3,0),1)', s('{1,2,3;4,0,0}')],
  ['WRAPCOLS', 'ARRAYTOTEXT(WRAPCOLS({1;2;3},5,0),1)', s('{1;2;3}')],
  ['WRAPCOLS', 'ARRAYTOTEXT(WRAPCOLS(5,1),1)', s('{5}')],
  ['WRAPCOLS', 'ARRAYTOTEXT(WRAPCOLS("a",4),1)', s('{"a"}')],
  // text as an array of numbers: no numbers in it
  ['F.TEST', 'F.TEST("abc",D!C1:C8)', e('DIV0')],
  ['F.TEST', 'F.TEST("5",D!C1:C8)', e('DIV0')],
  ['F.TEST', 'F.TEST(D!A3,D!C1:C8)', e('DIV0')],
  ['F.TEST', 'F.TEST({"a","b"},D!C1:C8)', e('DIV0')],
  ['FTEST', 'FTEST(D!B1:B8,"x")', e('DIV0')],
  ['MIRR', 'MIRR("abc",0.1,0.12)', e('DIV0')],
  ['MIRR', 'MIRR(TRUE,0.1,0.12)', e('DIV0')],
  ['T.TEST', 'T.TEST("",D!B1:B8,2,1)'],
  ['CORREL', 'CORREL("",D!B1:B8)'],
  ['IRR', 'IRR("")'],
  ['SUMPRODUCT', 'SUMPRODUCT("")'],
  // text or a boolean as XIRR's and XNPV's values or dates: #VALUE! before the sizes are compared
  ['XIRR', 'XIRR("5",D!AU1:AU5)', e('VALUE')],
  ['XIRR', 'XIRR(TRUE,D!AU1:AU5)', e('VALUE')],
  ['XIRR', 'XIRR(D!J1:J5,"x")', e('VALUE')],
  ['XIRR', 'XIRR(D!J1:J5,D!AU1)', e('NUM')],
  ['XIRR', 'XIRR(D!A3,D!AU1:AU5)'],
  ['XNPV', 'XNPV(0.09,"5",D!AU1:AU5)', e('VALUE')],
  ['XNPV', 'XNPV(0.09,D!J1:J5,"x")', e('VALUE')],
  ['XNPV', 'XNPV(0.09,TRUE,D!AU1:AU5)', e('VALUE')],
  ['XNPV', 'XNPV(0.09,D!J1:J5,D!AU1)', e('NUM')],
  ['XNPV', 'XNPV(0.1,-100,39448)'],
]

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [fn, f] of CASES) add(fn, fn, f)
  return out
}

/** The predicted answers by case id. */
export const PREDICTED: Readonly<Record<string, Val>> = Object.fromEntries(
  buildProbes().flatMap((p, i) => (CASES[i][2] ? [[p.id, CASES[i][2]]] : [])),
)
