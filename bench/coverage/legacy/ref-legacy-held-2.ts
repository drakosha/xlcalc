/**
 * Held out (`run.ts legacy ref-legacy-held-2`): arrays of references given to the functions that take them one
 * reference at a time since the spill set `ref-arrays-held-2` (RANK, ROW, COLUMN, AREAS, CELL, FORMULATEXT, ISFORMULA, the
 * D functions), and to AGGREGATE and INDEX, which answer #VALUE! for them, in formulas of the old kind without
 * Ctrl+Shift+Enter: alone, inside SUM and inside SUMPRODUCT. Predicted before Excel ran
 * (`ref-legacy-held-2.predicted.json`) by the rule `ref-legacy-held` measured for SUBTOTAL and COUNTIF: alone, the array
 * under the reference takes `@`; inside SUM or SUMPRODUCT it stays; AGGREGATE and INDEX, whose argument takes an array,
 * stay too; the `@` the converter puts before CELL, FORMULATEXT and ISFORMULA stays.
 */
import type { Probe } from '../probes.js'
import { probe } from './params.js'
import PREDICTED from './ref-legacy-held-2.predicted.json' with { type: 'json' }

export const FORMULAS = [
  'RANK(20,OFFSET(D!B1,{0;1},0,2))', 'SUM(RANK(20,OFFSET(D!B1,{0;1},0,2)))', 'RANK.EQ(20,OFFSET(D!B1,0,0,{2;3}),1)',
  'ROW(OFFSET(D!B1,{0;1;2},0))', 'SUM(ROW(OFFSET(D!B1,{0;1;2},0)))', 'COLUMN(OFFSET(D!B1,0,{0,1}))',
  'AREAS(OFFSET(D!B1,{0;1},0))', 'SUM(AREAS(OFFSET(D!B1,{0;1},0)))', 'CELL("address",OFFSET(D!B1,{0;1},0))',
  'FORMULATEXT(OFFSET(D!B1,{0;1},0))', 'ISFORMULA(OFFSET(D!B1,{0;1},0))', 'SUMPRODUCT(--ISFORMULA(OFFSET(D!B1,{0;1},0)))',
  'DSUM(OFFSET(D!A1,0,0,{2;3},2),2,D!A1:A2)', 'SUM(DSUM(OFFSET(D!A1,0,0,{2;3},2),2,D!A1:A2))', 'DCOUNT(D!A1:B3,2,OFFSET(D!A1,0,0,2,{1,2}))',
  'AGGREGATE(9,0,OFFSET(D!B1,{0;1},0,2))', 'INDEX(OFFSET(D!B1,{0;1},0,2),2)',
]

export function buildProbes(): Probe[] {
  return FORMULAS.map((f, i) => probe(`refs2#${i + 1}`, 'refs2', f.match(/^[A-Z.]+/)![0], f))
}

/** The text Excel 365 was predicted to show for each case (canonical, without `=`), committed before Excel ran. */
export const PREDICTED_SHOWN: Readonly<Record<string, string>> = PREDICTED

/**
 * The predictions Excel did not confirm (Excel 16.115 for Mac and 16.0 for Windows alike): ROW takes `@` inside SUM too, FORMULATEXT and ISFORMULA
 * of one cell take none before them, AGGREGATE's and INDEX's array of references takes `@` as SUBTOTAL's does.
 */
export const MISSED: readonly string[] = ['refs2#5', 'refs2#10', 'refs2#11', 'refs2#16', 'refs2#17']
