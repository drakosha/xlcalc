/**
 * Held out (`run.ts legacy ref-legacy-held`): arrays of references (OFFSET or INDIRECT over an array) given to SUBTOTAL,
 * COUNTIF, SUMIF, SUMIFS and COUNTBLANK in formulas of the old kind without Ctrl+Shift+Enter: alone, inside SUM and inside
 * SUMPRODUCT. The engine calls these functions once for each reference since the set `ref-arrays-held`; which text Excel
 * shows for the old formulas, and so whether they spill once converted, was not measured. Predicted before Excel ran
 * (`ref-legacy-held.predicted.json`): the converter's text, an array inside an argument that takes a reference left as
 * it is (as `SUM(N(OFFSET(…,{1;2},0)))` in `deref-held`), ROW over several cells with `@` outside an array.
 */
import type { Probe } from '../probes.js'
import { probe } from './params.js'
import PREDICTED from './ref-legacy-held.predicted.json' with { type: 'json' }

export const FORMULAS = [
  'SUBTOTAL(9,OFFSET(D!B1,{0;1;2},0))', 'SUM(SUBTOTAL(9,OFFSET(D!B1,{0;1;2},0)))', 'SUMPRODUCT(SUBTOTAL(3,OFFSET(D!B1,{0;1;2},0)))',
  'COUNTIF(INDIRECT({"D!B1:B2";"D!B2:B3"}),">5")', 'SUM(COUNTIF(INDIRECT({"D!B1:B2";"D!B2:B3"}),">5"))', 'SUMPRODUCT(COUNTIF(INDIRECT({"D!B1:B2";"D!B2:B3"}),">5"))',
  'SUMIF(OFFSET(D!B1,{0;1},0,2),">5")', 'SUM(SUMIF(OFFSET(D!B1,{0;1},0,2),">5"))', 'SUMPRODUCT(SUMIF(OFFSET(D!B1,{0;1},0,2),">5",OFFSET(D!C1,{0;1},0,2)))',
  'COUNTBLANK(OFFSET(D!D4,{0;1},0,2))', 'SUM(COUNTBLANK(OFFSET(D!D4,{0;1},0,2)))', 'SUM(SUBTOTAL(9,OFFSET(D!B1,ROW(D!B1:B3)-1,0)))',
  'SUMPRODUCT(SUBTOTAL(9,OFFSET(D!B1,ROW(D!B1:B3)-ROW(D!B1),0)))', 'SUM(COUNTIF(OFFSET(D!B1,{0;1;2},0),">5"))', 'MAX(SUBTOTAL(4,OFFSET(D!B1,0,{0,1},3)))',
  'SUMIFS(OFFSET(D!C1,{0;1},0,2),OFFSET(D!B1,{0;1},0,2),">5")', 'SUM(SUMIFS(OFFSET(D!C1,{0;1},0,2),OFFSET(D!B1,{0;1},0,2),">5"))',
  'SUM(SUBTOTAL(9,INDIRECT({"D!B1";"D!C1"})))', 'SUMPRODUCT(SUBTOTAL(9,OFFSET(D!B1,{0;1;2},0)))*1',
]

export function buildProbes(): Probe[] {
  return FORMULAS.map((f, i) => probe(`refs#${i + 1}`, 'refs', f.match(/^[A-Z]+/)![0], f))
}

/** The text Excel 365 was predicted to show for each case (canonical, without `=`), committed before Excel ran. */
export const PREDICTED_SHOWN: Readonly<Record<string, string>> = PREDICTED

/**
 * The predictions Excel did not confirm (Excel 16.115 for Mac and 16.0 for Windows alike): alone, SUBTOTAL, COUNTIF, SUMIF,
 * SUMIFS and COUNTBLANK do not evaluate an array under their reference, OFFSET's and INDIRECT's array takes `@` there.
 */
export const MISSED: readonly string[] = ['refs#1', 'refs#4', 'refs#7', 'refs#10', 'refs#16']
