/**
 * SORTBY with a by-array of one value and orders of one value each, held out. Excel answers
 * `SORTBY(D!B1:B8,D!C1:C8,1,5,1)` with #VALUE!, and `SORTBY(D!B1:B4,0,-1)` (the set `omit`) with the array as it is.
 * The rule predicted: a by-array of one value is one row and one column, so it fits a sort of rows when the array has
 * one row and a sort of columns when it has one column; the first by-array sets which, and every later one must fit
 * that sort too. An error given as a by-array is still the answer. Answers: `test/excel/cases/sortby-one-held.json`;
 * `PREDICTED`, the answers by that rule, was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { Probe } from '../probes.js'
import { collector } from './arrays/common.js'

const n = (v: number): Val => ({ t: 'n', v })
const s = (v: string): Val => ({ t: 's', v })
const e = (v: Extract<Val, { t: 'e' }>['v']): Val => ({ t: 'e', v })

/** Function, formula and the answer predicted before Excel ran (D!B1:B8 4, 8, 15, 16, 23, 42, −3.5, 10.25 by D!C1:C8). */
const CASES: [string, string, Val][] = [
  // a column of eight sorted by a column, then one value: a literal, an order left out, a cell, an array of one, text, as the third
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C8,1,5,1),1)', e('VALUE')],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C8,1,5,),1)', e('VALUE')],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C8,-1,D!E1,1),1)', e('VALUE')],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C8,1,{5},1),1)', e('VALUE')],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C8,1,"x",1),1)', e('VALUE')],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C8,1,D!E1:E8,1,5,1),1)', e('VALUE')],
  // one value first: a column after it does not fit, another value does; a table of two columns takes none
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,5,1,D!C1:C8,1),1)', e('VALUE')],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,5,1,7,-1),1)', s('{4;8;15;16;23;42;-3.5;10.25}')],
  ['SORTBY', 'ROWS(SORTBY(D!B1:B8,5))', n(8)],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:C8,5,1),1)', e('VALUE')],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:C8,,-1),1)', e('VALUE')],
  // a row: sorted by a row, then one value; one value alone; one value, then a row
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!L8:Q8,D!L10:Q10,-1,5,1),1)', e('VALUE')],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!L8:Q8,5,1),1)', s('{1,2,3,5,8,13}')],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!L8:Q8,5,1,D!L10:Q10,-1),1)', e('VALUE')],
  // one cell by one cell and one value; an error as the by-array; a by-array as long as the array
  ['SORTBY', 'SORTBY(D!B1,D!C1,1,5,1)', n(4)],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C8,1,1/0,1),1)', e('DIV0')],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C8,1,D!B1:B8*0,1),1)', s('{-3.5;4;8;10.25;15;16;23;42}')],
  // orders by position: the call at each position
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C8,{1;-1},5,1),1)', s('{#VALUE!;#VALUE!}')],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,5,{1;-1}),1)', s('{4;}')],
]

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [fn, f] of CASES) add(fn, fn, f)
  return out
}

/** The predicted answers by case id. */
export const PREDICTED: Readonly<Record<string, Val>> = Object.fromEntries(buildProbes().map((p, i) => [p.id, CASES[i][2]]))

/** The predictions Excel did not confirm, the Mac and Windows alike: none. */
export const MISSED: readonly string[] = []
