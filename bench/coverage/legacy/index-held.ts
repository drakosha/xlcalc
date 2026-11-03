/**
 * Held out from `index-calls` and `offset-calls` (`run.ts legacy index-held`): INDEX with MATCH in more shapes and
 * around more calls, MATCH changed by an operator, OFFSET of one cell's size in more ways. The text Excel shows is the
 * converter's, predicted before Excel ran the set (`index-held.predicted.json`).
 */
import type { Probe } from '../probes.js'
import { probe } from './params.js'
import PREDICTED from './index-held.predicted.json' with { type: 'json' }

const M8 = 'MATCH(8,D!B1:B8,0)'

const FORMULAS = [
  'INDEX(D!B1:C3,(2),1)', 'INDEX(D!$B$1:$B$8,MATCH(8,D!$B$1:$B$8,0))', 'INDEX(D!B1:C3,MATCH("Чай",D!F1:F7,0)-4,2)', `INDEX(D!B1:C3,2,${M8})`,
  `INDEX(D!B1:C8,${M8},MATCH(7,D!C1:C8,0),1)`, `INDEX(D!B1:B8,${M8}*1)`, `INDEX(D!B1:B8,${M8}+0)`, `INDEX(D!B1:B8,+${M8})`, 'INDEX(D!B1:B8,LOOKUP(2,D!L1:L6))',
  'INDEX(D!B1:B8,MATCH(8,D!B1:B8))', 'IFERROR(INDEX(D!B1:B8,MATCH(99,D!B1:B8,0)),"")', `INDEX(D!B1:B8,${M8}):D!B8`, `VLOOKUP(INDEX(D!L1:L6,${M8}),D!L1:N6,2,FALSE)`,
  'INDEX(D!B1:B8,MATCH(D!A2,D!C1:C8,0))', 'INDEX(D!B1:B8,MATCH(1,(D!B1:B8>5)*(D!C1:C8>5),0))', 'INDEX(D!B1:C3,1,1)', `INDEX(D!B1:C3,${M8},MATCH(D!A12,D!C1:C8,0))`,
  `INDEX(D!B1:B8,MATCH(8,D!B1:B8,0),1)`, `INDEX(D!B1:C3,${M8},1)*INDEX(D!B1:C3,1,${M8})`, `IF(D!A4,INDEX(D!B1:B8,${M8}),0)`, `INDEX(D!B1:B8,INDEX(D!L1:L6,${M8}))`,
  `INDEX(D!B1:B8,MAX(${M8},1))`, `ROWS(INDEX(D!B1:C3,${M8},0))`, `INDEX(D!B1:B8,${M8})>5`, `TEXT(INDEX(D!B1:B8,${M8}),"0")`,
  `OFFSET(D!B1,${M8},1,1)`, 'OFFSET(D!B1:C1,0,0,1,1)', 'OFFSET(D!B1,0,0,1,D!G1)', 'OFFSET(D!A1,ROW()-1,0)', `SUM(OFFSET(D!B1,0,0,${M8}))`, 'OFFSET(D!B1:B1,1,0)',
  'OFFSET(D!B1,1,0,1,)', 'OFFSET(INDIRECT("D!B1"),1,0)', 'OFFSET(D!B1,1,0)*OFFSET(D!C1,1,0)', 'IFERROR(OFFSET(D!B1,9,0),0)', `OFFSET(INDEX(D!B1:B8,${M8}),0,1)`,
  'OFFSET(D!B1:C3,1,1,1,1)+1', 'OFFSET(D!B1,1,0,1,1.5)',
]

export function buildProbes(): Probe[] {
  return FORMULAS.map((f, i) => probe(`held#${i + 1}`, 'held', f.match(/^[A-Z]+/)![0], f))
}

/** The text Excel 365 was predicted to show for each case (canonical, without `=`), committed before Excel ran. */
export const PREDICTED_SHOWN: Readonly<Record<string, string>> = PREDICTED

/** The predictions Excel did not confirm. */
export const MISSED: readonly string[] = []
