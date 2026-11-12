/**
 * Held out (`run.ts legacy deref-held`): the dereferencing idioms and the classic array tricks of formulas of the old
 * kind without Ctrl+Shift+Enter. 0.89.1 put `@` inside N where N stands in a lookup's key or place
 * (`SUM(INDEX(B1:B3,N(@IF(1,{2;3}))))`, 0.89.0 left the text as it was); Excel had not shown its text for that. The
 * predicted text (`deref-held.predicted.json`, committed before Excel ran the set) is the converter's of 0.89.1: N's
 * array takes `@` where one value is wanted, as `N(@{1,2})` at the top (`calls`) and a lookup's place inside SUM
 * (`array-held`) show. OFFSET, INDIRECT, SUBTOTAL, VLOOKUP's column, LOOKUP and COUNTIF are controls predicted unchanged.
 */
import type { Probe } from '../probes.js'
import { probe } from './params.js'
import PREDICTED from './deref-held.predicted.json' with { type: 'json' }

export const FORMULAS = [
  'SUM(INDEX(D!B1:B3,N(IF(1,{2;3}))))', 'SUMPRODUCT(INDEX(D!B1:B3,N(IF(1,{2;3}))))', 'SUM(XLOOKUP(N(IF(1,{1;2})),D!E1:E8,D!B1:B8))',
  'SUMPRODUCT(XLOOKUP(N(IF(1,{1;2})),D!E1:E8,D!B1:B8))', 'MAX(INDEX(D!B1:B3,N(IF(1,{1;3}))))', 'SUM(INDEX(D!B1:B3,N(IF(1,D!E1:E2))))',
  'SUM(INDEX(D!B1:B3,N({2;3})))', 'SUMPRODUCT(N(OFFSET(D!B1,{1;2},0)))', 'SUM(N(OFFSET(D!B1,{1;2},0)))', 'SUMPRODUCT(N(INDIRECT({"D!B2";"D!B3"})))',
  'SUMPRODUCT(SUBTOTAL(9,OFFSET(D!B1,{0;1;2},0)))', 'SUM(VLOOKUP(4,D!B1:D3,{2,3},FALSE))', 'LOOKUP(2,1/(D!B1:B3>5),D!C1:C3)',
  'INDEX(D!C1:C3,MATCH(1,INDEX((D!B1:B3=8)*(D!C1:C3>5),0),0))', 'SUM(COUNTIF(D!B1:B3,{8,15}))', 'SUM(N({2;3}))', 'SUMPRODUCT(N({2;3}))',
]

export function buildProbes(): Probe[] {
  return FORMULAS.map((f, i) => probe(`deref#${i + 1}`, 'deref', f.match(/^[A-Z]+/)![0], f))
}

/** The text Excel 365 was predicted to show for each case (canonical, without `=`), committed before Excel ran. */
export const PREDICTED_SHOWN: Readonly<Record<string, string>> = PREDICTED

/** The predictions Excel did not confirm. */
export const MISSED: readonly string[] = []
