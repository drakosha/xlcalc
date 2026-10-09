/**
 * What the rules of `known-held` left unmeasured, held out: SORTBY with two orders of one shape, with valid orders and
 * a by-array of another shape, with an order of one value beside one of two; LOOKUP reading on past the sheet's last
 * column and last row (sheet Lk holds values at its edges, so a blank and a cell read tell apart), and with a result
 * that a function answers. Answers: `test/excel/cases/lookup-held.json`; `PREDICTED`, the engine's answers, was
 * committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'

/** Lk: text at the end of row 1 (XFB1:XFD1) and of columns A to C (rows 1048575 and 1048576). */
export const SHEETS: SheetCells = {
  Lk: { XFB1: 'xfb', XFC1: 'xfc', XFD1: 'xfd', A1048575: 'a5', A1048576: 'a6', B1048576: 'b6', C1048576: 'c6', B2: 'b2', B3: 'b3', B4: 'b4', C2: 'c2' },
}

const s = (v: string): Val => ({ t: 's', v })
const e = (v: Extract<Val, { t: 'e' }>['v']): Val => ({ t: 'e', v })

/** Function, formula and the answer predicted before Excel ran (4.19 falls in D!L3, 5 in D!L4 and D!O8). */
const CASES: [string, string, Val][] = [
  // SORTBY: two orders of one shape, by position
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C8,{1;-1},D!B1:B8,{1;-1}),1)', s('{-3.5;}')],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C8,{1;-1},D!B1:B8,{-1;1}),1)', s('{-3.5;}')],
  // valid orders, a by-array of another shape than the array
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C7,{1;-1}),1)', s('{#VALUE!;}')],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,TRANSPOSE(D!C1:C8),{1;-1}),1)', s('{#VALUE!;}')],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C8,{1;-1},D!B1:B6,{1;-1}),1)', s('{#VALUE!;}')],
  // an order of one value beside one of two
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C8,{1},D!B1:B8,{1;-1}),1)', s('{-3.5;#VALUE!}')],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C8,{-1},D!B1:B8,{1,-1}),1)', s('{42,#VALUE!}')],
  // LOOKUP past the sheet's last column: a cell is read as a row
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6,Lk!XFB1)&""', s('xfd')],
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6,Lk!XFD1)&""', s('')],
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6,Lk!XFC1:XFD1)&""', s('')],
  // past the sheet's last row: a column read down, a cell or a row of the last row read right
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6,Lk!A1048575:A1048576)&""', s('')],
  ['LOOKUP', 'LOOKUP(5,D!L8:Q8,Lk!A1048575:A1048576)&""', s('')],
  ['LOOKUP', 'LOOKUP(2,D!L1:L6,Lk!A1048576)&""', s('b6')],
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6,Lk!A1048576:B1048576)&""', s('c6')],
  // a result that a function answers
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6,IF(TRUE,Lk!B2:B3))&""', e('REF')],
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6,CHOOSE(1,Lk!B2:B3))&""', e('REF')],
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6,INDEX(Lk!B2:B3,0,1))&""', e('REF')],
  ['LOOKUP', 'LOOKUP(2,D!L1:L6,IF(TRUE,Lk!B2))&""', e('REF')],
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6,IF(TRUE,Lk!XFD1))&""', s('')],
]

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [fn, f] of CASES) add(fn, fn, f)
  return out
}

/** The predicted answers by case id. */
export const PREDICTED: Readonly<Record<string, Val>> = Object.fromEntries(buildProbes().map((p, i) => [p.id, CASES[i][2]]))
