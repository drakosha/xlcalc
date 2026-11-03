/**
 * Held out after the set `index-held` (`run.ts legacy index-held-2`): OFFSET's reference of one cell against a range
 * of one cell with each kind of size, and INDEX with MATCH in more calls. The text Excel shows is the converter's,
 * predicted before Excel ran the set (`index-held-2.predicted.json`).
 */
import type { Probe } from '../probes.js'
import { probe } from './params.js'
import PREDICTED from './index-held-2.predicted.json' with { type: 'json' }

const M8 = 'MATCH(8,D!B1:B8,0)'

const FORMULAS = [
  'OFFSET(D!B1:B1,0,0,1,1)', 'OFFSET(D!B1:B1,1,0,1)', 'OFFSET(D!B1:C1,1,0,,1)', 'OFFSET((D!B1),1,0,1)', 'OFFSET(D!B1,1,0,,)', `OFFSET(D!$B$1,${M8},0)`,
  'OFFSET(D!B1,1,0,1,1)&OFFSET(D!B1:B3,1,0)', 'ABS(OFFSET(D!B1:B3,1,0,1,1))', `INDEX(D!B1:B8,${M8})*INDEX(D!C1:C8,${M8})`, `SUMPRODUCT(INDEX(D!B1:C3,${M8},0))`,
  `INDEX(D!B1:C3,MIN(${M8},2),1)`, `INDEX(D!B1:C3,1,(((${M8}))))`, `INDEX(D!B:C,${M8},2)`, `INDEX(D!B1:B8,MATCH(8,{4,8,15},0))`, `INDEX(D!B1:B8,MATCH(8,D!B1:B8,-1))`,
]

export function buildProbes(): Probe[] {
  return FORMULAS.map((f, i) => probe(`held#${i + 1}`, 'held', f.match(/^[A-Z]+/)![0], f))
}

/** The text Excel 365 was predicted to show for each case (canonical, without `=`), committed before Excel ran. */
export const PREDICTED_SHOWN: Readonly<Record<string, string>> = PREDICTED

/** The predictions Excel did not confirm. */
export const MISSED: readonly string[] = []
