/**
 * Held out after the set `groupby`, its rules predicted: headers left out (#VALUE! where text stands above a number in
 * the values, none elsewhere), headers made up and shown, a value spelled as the row that sorts first, the grand total
 * alone by default and `Grand Total` beside subtotals, a column of functions, filters, sorting, PERCENTOF and a LAMBDA of
 * two parameters, the table relationship, and PIVOTBY with subtotal columns, its headers, several functions and the
 * totals PERCENTOF is relative to. Over the table of the set `groupby` (sheet G) and array constants. Answers:
 * `test/excel/cases/groupby-held.json`; `PREDICTED`, the engine's answers (`sets/predict.ts groupby-held`), was committed
 * before Excel ran the set.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe } from '../probes.js'
import { xl } from './arrays/common.js'
import { xlpm } from './let/xlpm.js'
import { SHEETS as TABLE } from './groupby.js'

export const SHEETS = TABLE

const t = (f: string) => `ARRAYTOTEXT(${f},1)`
const [reg, regH, rp, year, sales, salesH, qtyH, keep] = ['G!A2:A11', 'G!A1:A11', 'G!A2:B11', 'G!C2:C11', 'G!D2:D11', 'G!D1:D11', 'G!E1:E11', 'G!F2:F11']
/** Two fields whose second is spelled two ways in one group and one way in another. */
const twoWays = '{"x","A";"x","a";"y","a";"y","B"}'

const CASES: Record<string, string[]> = {
  GROUPBY: [
    t(`GROUPBY(${regH},${qtyH},SUM)`),
    t(`GROUPBY(G!B2:B11,${reg},COUNTA)`),
    t(`GROUPBY(G!B1:B11,G!A1:A11,COUNTA)`),
    t(`GROUPBY(${reg},${sales},SUM,1)`),
    t(`GROUPBY(${rp},${sales},SUM,2)`),
    t(`GROUPBY(${rp},${sales},SUM,3,,,,)`),
    t(`GROUPBY(G!A1:B11,${salesH},SUM,3,2)`),
    t(`GROUPBY(${twoWays},{1;2;3;4},SUM)`),
    t(`GROUPBY(${twoWays},{1;2;3;4},SUM,0,1,,,1)`),
    t(`GROUPBY(${rp},${sales},SUM,0,3)`),
    t(`GROUPBY(${reg},${sales},SUM,0,-1)`),
    t(`GROUPBY(${reg},${sales},SUM,0,1,2)`),
    t(`GROUPBY(${rp},${sales},SUM,0,2,-2)`),
    t(`GROUPBY(${rp},${sales},SUM,0,2,3)`),
    t(`GROUPBY(${reg},${sales},VSTACK(SUM,MAX),3)`),
    t(`GROUPBY(${reg},${sales},SUM,,,,TRANSPOSE(${keep}))`),
    t(`GROUPBY(${reg},${sales},SUM,,,,G!F1:F11)`),
    t(`GROUPBY(${reg},${sales},PERCENTOF,,,,${keep})`),
    t(`GROUPBY(${reg},${sales},LAMBDA(x,y,SUM(x)/SUM(y)))`),
    t(`GROUPBY(${reg},${sales},LAMBDA(x,ROWS(x)))`),
    t(`GROUPBY(${reg},${sales},LAMBDA(x,INDEX(x,1)))`),
    t('GROUPBY({1;2;1;2},{"a";"b";"c";"d"},ARRAYTOTEXT)'),
    t(`GROUPBY(G!A2:A13,G!D2:D13,SUM)`),
    t(`GROUPBY(${reg},${sales},SUM,,,,${keep}*1)`),
    t(`GROUPBY(${reg},${sales},HSTACK(SUM,LAMBDA(x,MIN(x))))`),
    t(`GROUPBY(${reg},${sales},SUM,,,{1})`),
    t(`GROUPBY(${reg},${sales},SUM,,,3)`),
    t(`GROUPBY(${reg},G!D2:E11,SUM,,,-3)`),
    t(`GROUPBY(${reg},${sales},SUM,,,,,2)`),
    t(`GROUPBY(${reg},${sales},SUM,1.9)`),
  ],
  PIVOTBY: [
    t(`PIVOTBY(${reg},G!B2:C11,${sales},SUM,0,0,,2)`),
    t(`PIVOTBY(${reg},G!B2:C11,${sales},SUM,0,0,,-2)`),
    t(`PIVOTBY(${rp},${year},${sales},SUM,0,2)`),
    t(`PIVOTBY(${reg},${year},${sales},SUM,2)`),
    t(`PIVOTBY(G!A1:A11,G!B1:C11,${salesH},SUM,3,0,,0)`),
    t(`PIVOTBY(${reg},${year},${sales},VSTACK(SUM,COUNT))`),
    t(`PIVOTBY(${reg},${year},${sales},HSTACK(SUM,COUNT),3)`),
    t(`PIVOTBY(${rp},G!B2:B11,${sales},PERCENTOF,0,2,,,,,3)`),
    t(`PIVOTBY(${rp},${year},${sales},PERCENTOF,0,2,,,,,4)`),
    t(`PIVOTBY(${reg},${year},${sales},SUM,,,,,2)`),
    t(`PIVOTBY(${reg},${year},${sales},SUM,,,,,-3)`),
    t(`PIVOTBY(${reg},${year},${sales},LAMBDA(x,ROWS(x)))`),
    t(`PIVOTBY(${reg},${year},${sales},SUM,0,-1,,-1)`),
    t(`PIVOTBY(${reg},${year},${sales},SUM,,,,,,,5)`),
  ],
}

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  for (const [fn, formulas] of Object.entries(CASES)) formulas.forEach((f, i) => out.push({ id: `${fn}#${i + 1}`, group: fn, fn, formula: xl(xlpm(f)) }))
  return out
}

/** The engine's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./groupby-held.predicted.json', import.meta.url), 'utf8'))

/**
 * The predictions Excel for Mac did not confirm: headers left out are found where the first row field and the first
 * value are text and every value under them a number or blank (#VALUE! where another is not); a value is spelled as the
 * row that sorts first by the sort order given, in the table relationship too; the function gets a group's values in the
 * data's order; a LAMBDA's label is `CUSTOM`; a number not whole is no option; PIVOTBY joins the column fields' headers in
 * one cell and takes a column of functions down its rows; `LAMBDA(x,INDEX(x,1))` makes the summary #CALC!.
 */
export const MISSED: readonly string[] = ['GROUPBY#1', 'GROUPBY#2', 'GROUPBY#3', 'GROUPBY#9', 'GROUPBY#13', 'GROUPBY#21', 'GROUPBY#22', 'GROUPBY#25', 'GROUPBY#30', 'PIVOTBY#5', 'PIVOTBY#6', 'PIVOTBY#7']
