/**
 * LOOKUP beside the cells it may read on to, held out: a formula next to the column a spill fills (an array of 3 and
 * of 5, one that cannot spill), next to a column and a row a function answers, and in the row of a row result, past
 * the cells it may read and among them; a column INDEX takes from a larger rectangle. Each LOOKUP stands on the set's
 * sheets and a case reads its cell. Answers: `test/excel/cases/lookup-own-held.json`; `PREDICTED`, the engine's answers,
 * was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'
import { s } from '../vals.js'

const keys = Object.fromEntries(Array.from({ length: 6 }, (_, i) => [`A${i + 1}`, i + 1]))
const lookup = (formula: string) => ({ formula, dynamic: true as const })
const spill = (rows: number) => ({ ...keys, D1: rows, B1: lookup('_xlfn.SEQUENCE(D1)*10'), C1: lookup('LOOKUP(5,A1:A6,B1#)&""') })
const row = { ...keys, B10: 'b', C10: 'c', D10: 'd', F10: lookup('UPPER("f")') }

/**
 * Ls, Lt: B1 spills 3 or 5 values beside C1's LOOKUP, which reads B5; Ls also C2 and C3 over E1:E2 (E4, E5 past it);
 * Lv: B5 holds a value, B1 cannot spill. Lr, Lq, Lp: B10:D10 as a row, F10 read on to; the LOOKUP past what it may read
 * (I10) or among it (E10). Lu: INDEX takes the column C1:C3 of B1:D3, C5 read on to, the LOOKUP right of the rectangle.
 */
export const SHEETS: SheetCells = {
  Ls: { ...spill(3), E1: 'e1', E2: 'e2', E4: 'x', E5: 7, C2: lookup('LOOKUP(5,A1:A6,IF(TRUE,E1:E2))&""'), C3: lookup('LOOKUP(4,A1:A6,(E1:E2))&""') },
  Lt: spill(5),
  Lv: { ...spill(5), B5: 99 },
  Lr: { ...row, I10: lookup('LOOKUP(5,A1:A6,IF(TRUE,B10:D10))&""') },
  Lq: { ...row, E10: lookup('LOOKUP(5,A1:A6,IF(TRUE,B10:D10))&""') },
  Lp: { ...row, E10: lookup('LOOKUP(5,A1:A6,B10:D10)&""') },
  Lu: { ...keys, B1: 'b1', C1: 'c1', D1: 'd1', B2: 'b2', C2: 'c2', D2: 'd2', B3: 'b3', C3: 'c3', D3: 'd3', C5: 'c5', F2: lookup('LOOKUP(5,A1:A6,INDEX(B1:D3,0,2))&""') },
}

/** The engine's #CYCLE!, which Excel has no code for: where the engine counts the LOOKUP's own cell among those it may read. */
const cycle = { t: 'e', v: 'CYCLE' } as unknown as Val

/** The cell read and the answer predicted before Excel ran. */
const CASES: [string, Val][] = [
  ['Ls!C1', s('')],
  ['Ls!C2', s('7')],
  ['Ls!C3', s('x')],
  ['Lt!C1', s('50')],
  ['Lv!C1', cycle],
  ['Lr!I10', s('F')],
  ['Lq!E10', cycle],
  ['Lp!E10', cycle],
  ['Lu!F2', cycle],
]

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [f] of CASES) add('LOOKUP', 'LOOKUP', f)
  return out
}

/** The predicted answers by case id. */
export const PREDICTED: Readonly<Record<string, Val>> = Object.fromEntries(buildProbes().map((p, i) => [p.id, CASES[i][1]]))

/**
 * The predictions Excel did not confirm (the Mac and Windows alike), all the engine's #CYCLE!: Excel follows the cells
 * LOOKUP reads and finds no cycle; the engine counts every cell it may read, its own among them (`lookupReadZone`).
 */
export const MISSED: readonly string[] = ['LOOKUP#5', 'LOOKUP#7', 'LOOKUP#8', 'LOOKUP#9']
