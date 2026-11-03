/**
 * OFFSET in formulas of the old kind (`run.ts legacy offset-calls`): the set `index-calls` showed `OFFSET(D!B1,x,0)`
 * and `OFFSET(D!B1,x,0,1,1)` without `@`, a range or a computed height with it. Here one cell by the reference and
 * the size (height and width the literal 1 or left out, the reference one cell, any shift) against a range, a size
 * of another literal, an expression or a cell, and OFFSET inside other calls and operators. The text Excel shows was
 * predicted before Excel ran the set (`offset-calls.predicted.json`).
 */
import type { Probe } from '../probes.js'
import { probe } from './params.js'
import PREDICTED from './offset-calls.predicted.json' with { type: 'json' }

const M8 = 'MATCH(8,D!B1:B8,0)'

const FORMULAS = [
  'OFFSET(D!B1,1,0)', `OFFSET(D!B1,${M8}-1,0)`, 'OFFSET(D!B1,D!A12,0)', 'OFFSET(D!B1,0,D!A12)', 'OFFSET(D!B1,1,1,1,1)', 'OFFSET(D!B1,1,0,1)',
  'OFFSET(D!B1,1,0,,1)', 'OFFSET(D!B1,1,0,2)', 'OFFSET(D!B1,1,0,1,2)', `OFFSET(D!B1,1,0,${M8})`, 'OFFSET(D!B1,1,0,1+0)', 'OFFSET(D!B1,1,0,D!G1)',
  'OFFSET(D!B1:B3,1,0)', 'OFFSET(D!B1:B3,1,0,1)', 'OFFSET(D!B1:C1,1,0,1)', 'OFFSET(D!B1:C3,1,0,1,1)', 'OFFSET(D!B:B,1,0,1)', 'SUM(OFFSET(D!B1,0,0,3))',
  'SUM(OFFSET(D!B1,1,0))', 'OFFSET(D!B1,1,0)+1', 'OFFSET(D!B1,1,0)&""', 'ABS(OFFSET(D!B1,1,0))', 'OFFSET(D!B1,1,0,2)+1', `OFFSET(D!B1,${M8},${M8})`,
  'OFFSET(D!B1,1.5,0)', 'OFFSET(D!B1,{1,2},0)', 'OFFSET(D!B1,1,0,1.5)', 'OFFSET((D!B1),1,0)', 'OFFSET(INDEX(D!B1:B8,2),1,0)', 'OFFSET(D!B1,-1,0)',
  'OFFSET(D!B1,1,0,-1)', 'OFFSET(D!B1,1,0,0)', 'IF(TRUE,OFFSET(D!B1,1,0))', 'INDEX(OFFSET(D!B1,0,0,3),2)', 'ROWS(OFFSET(D!B1,0,0,3))', 'OFFSET(D!B1,1,0,1,1)*2',
  'IFERROR(OFFSET(D!B1,0,0,3),0)', 'IFERROR(OFFSET(D!B1,1,0),0)', 'OFFSET(D!B1,1,0,TRUE)', 'OFFSET(D!B1,,)',
]

export function buildProbes(): Probe[] {
  return FORMULAS.map((f, i) => probe(`offset#${i + 1}`, 'offset', 'OFFSET', f))
}

/** The text Excel 365 was predicted to show for each case (canonical, without `=`), committed before Excel ran. */
export const PREDICTED_SHOWN: Readonly<Record<string, string>> = PREDICTED

/** The predictions Excel did not confirm: a size left empty or out over a range of one column or row is not one cell to Excel. */
export const MISSED: readonly string[] = ['offset#7', 'offset#14', 'offset#17']
