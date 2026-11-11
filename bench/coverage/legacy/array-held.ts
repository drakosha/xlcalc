/**
 * Held out (`run.ts legacy array-held`): an array, a constant or a call's answer, in an argument Excel stops array
 * evaluation at (class `V` of `params.ts`: XLOOKUP's key, VLOOKUP's, INDEX's places, IF's condition, CHOOSE's index,
 * TRANSPOSE) inside SUM, SUMPRODUCT and AVERAGE, with arguments of class `v` as controls. Excel showed
 * `SUM(XLOOKUP(@{1;2},…))` for an old formula (11.11.2026); the converter put `@` there at the top level only. The
 * text Excel shows is the converter's after the change, predicted before Excel ran the set (`array-held.predicted.json`).
 */
import type { Probe } from '../probes.js'
import { probe } from './params.js'
import PREDICTED from './array-held.predicted.json' with { type: 'json' }

const LOOK = 'D!B1:B8,D!C1:C8'

export const FORMULAS = [
  `SUM(XLOOKUP({1;2},${LOOK}))`, `SUMPRODUCT(XLOOKUP({1;2},${LOOK}))`, `AVERAGE(XLOOKUP({1;2},${LOOK}))`, `SUM(XLOOKUP({1;2}+0,${LOOK}))`,
  `SUM(XLOOKUP(ABS({-1;-2}),${LOOK}))`, `SUM(XLOOKUP(SEQUENCE(2),${LOOK}))`, `SUM(XLOOKUP(1,${LOOK},,{0;1}))`, `SUM(IFERROR(XLOOKUP({1;2},${LOOK}),0))`,
  'SUM(LET(k,{1;2},XLOOKUP(k,D!B1:B8,D!C1:C8)))', 'SUM(VLOOKUP({1;2},D!B1:C8,2,FALSE))', 'SUM(HLOOKUP({1,2},D!B1:C8,2,FALSE))',
  'SUM(INDEX(D!B1:B8,{2;3}))', 'SUM(INDEX(D!B1:C8,1,{1,2}))', 'COUNT(INDEX(D!B1:B8,SEQUENCE(2)))', 'SUM(IF({TRUE;FALSE},1,2))',
  'SUM(CHOOSE({1;2},D!B1,D!B2))', 'SUM(TRANSPOSE({1,2}))', 'SUM(ABS({-1;-2}))', 'SUM(MOD({5;7},3))', 'SUM(XMATCH({1;2},D!B1:B8))',
]

export function buildProbes(): Probe[] {
  return FORMULAS.map((f, i) => probe(`array#${i + 1}`, 'array', f.match(/^[A-Z]+/)![0], f))
}

/** The text Excel 365 was predicted to show for each case (canonical, without `=`), committed before Excel ran. */
export const PREDICTED_SHOWN: Readonly<Record<string, string>> = PREDICTED

/** The predictions Excel did not confirm. */
export const MISSED: readonly string[] = ['array#15', 'array#16', 'array#17']
