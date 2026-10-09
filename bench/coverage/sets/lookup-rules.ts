/**
 * The rules that the misses of `lookup-held` led to, measured after it: SORTBY with arrays of orders answers, at each
 * position in rows, the error of its call with the orders there (an order of one row or column repeats, past a
 * shorter one #N/A), the first value of the first call that sorts, and after it a blank, or #VALUE! when it has more
 * than one by-array; LOOKUP reads on past the end of a reference that a spill, a function or parentheses answer.
 * Answers: `test/excel/cases/lookup-rules.json`.
 */
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'

/** Lk: B2:B4 and C2:C3, G3 beside the end of E2's array of two (E2:E3), E4 below it. */
export const SHEETS: SheetCells = {
  Lk: { B2: 'b2', B3: 'b3', B4: 'b4', C2: 'c2', C3: 'c3', E2: { formula: '_xlfn.SEQUENCE(2)', dynamic: true }, E4: 'e4', G3: 'g3' },
}

const CASES: [string, string][] = [
  // the first position that sorts need not be the first one
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!M1:M6,D!N1:N6,{2;-1}),1)'],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!M1:M6,D!N1:N6,{#N/A;-1;1}),1)'],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!M1:M6,D!N1:N6,{2,1;-1,1}),1)'],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C8,{1;-1},D!B1:B8,{#N/A;1}),1)'],
  // more than one by-array: #VALUE! after the first that sorts, a scalar order too; past a shorter order #N/A
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C8,{1;-1},D!B1:B8,1),1)'],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C8,{1;-1},D!B1:B8,{1;-1;1}),1)'],
  // an error given as a by-array comes before an order's error, which comes before a by-array of another size
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,1/0,{1;#N/A}),1)'],
  ['SORTBY', 'ARRAYTOTEXT(SORTBY(D!B1:B8,D!C1:C7,{1;#N/A}),1)'],
  // LOOKUP past the end of a spill, of a cell a function takes from it, of parentheses, of a range operator
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6,Lk!E2#)&""'],
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6,TAKE(Lk!E2#,-1))&""'],
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6,(Lk!B2:B3))&""'],
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6,Lk!B2:Lk!B3)&""'],
  ['LOOKUP', 'LOOKUP(2,D!L1:L6,INDEX(Lk!B2:B3,2))&""'],
  ['LOOKUP', 'LOOKUP(4.19,D!L1:L6+0,IF(TRUE,Lk!B2:B3))&""'],
]

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [fn, f] of CASES) add(fn, fn, f)
  return out
}
