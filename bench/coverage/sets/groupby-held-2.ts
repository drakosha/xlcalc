/**
 * Held out after `groupby-held`, its rules predicted, and what its one unexplained miss asks: which answers of a
 * function GROUPBY and MAP take as one value (`INDEX(x,1)` made the summary #CALC!: an array of one, a reference, a
 * value read through `@`); headers left out over blanks and over two value columns; the spelling of a value by a sort
 * order on a value column; fractional sort orders, depths and relationships; PIVOTBY's joined column headers over
 * generated ones and a column of functions with headers. Answers: `test/excel/cases/groupby-held-2.json`; `PREDICTED`,
 * the engine's answers (`sets/predict.ts groupby-held-2`), was committed before Excel ran the set.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe } from '../probes.js'
import { xl } from './arrays/common.js'
import { xlpm } from './let/xlpm.js'
import { SHEETS as TABLE } from './groupby.js'

export const SHEETS = { ...TABLE, H: { A1: 'Key', B1: 'Amount', C1: 'Extra', A2: 'p', B2: 1, C2: 5, A3: 'q', C3: 6, A4: 'p', B4: 3, C4: 7 } }

const t = (f: string) => `ARRAYTOTEXT(${f},1)`
const [reg, rp, year, sales] = ['G!A2:A11', 'G!A2:B11', 'G!C2:C11', 'G!D2:D11']

const CASES: Record<string, string[]> = {
  GROUPBY: [
    t(`GROUPBY(${reg},${sales},LAMBDA(x,INDEX(x,1,1)))`),
    t(`GROUPBY(${reg},${sales},LAMBDA(x,TAKE(x,1)))`),
    t(`GROUPBY(${reg},${sales},LAMBDA(x,@x))`),
    t(`GROUPBY(${reg},${sales},LAMBDA(x,SUM(INDEX(x,1))))`),
    t(`GROUPBY(${reg},${sales},LAMBDA(x,{5}))`),
    t(`GROUPBY(${reg},${sales},LAMBDA(x,ROWS(INDEX(x,1))))`),
    t('GROUPBY(H!A1:A4,H!B1:B4,SUM)'),
    t('GROUPBY(H!A1:A4,H!B1:C4,SUM)'),
    t('GROUPBY(H!A1:A4,H!C1:C4,SUM)'),
    t(`GROUPBY(${rp},${sales},SUM,0,1,-3)`),
    t(`GROUPBY(${reg},${sales},SUM,,,1.5)`),
    t(`GROUPBY(${reg},${sales},SUM,0,1.5)`),
    t(`GROUPBY(${rp},${sales},SUM,0,1,,,0.5)`),
    t(`GROUPBY(${reg},${sales},SUM,,,-2,,)`),
    t(`GROUPBY(${reg},${sales},LAMBDA(a,b,ROWS(b)))`),
    t(`GROUPBY(${reg},${sales},HSTACK(SUM,PERCENTOF),3)`),
  ],
  MAP: [t('MAP({1;2},LAMBDA(x,TAKE({7;8},1)))'), t('MAP({1;2},LAMBDA(x,INDEX({7;8},1)))'), t('MAP({1;2},LAMBDA(x,G!D2))')],
  PIVOTBY: [
    t(`PIVOTBY(${reg},G!B2:C11,${sales},SUM,2,0,,0)`),
    t(`PIVOTBY(${reg},${year},${sales},VSTACK(SUM,COUNT),3)`),
    t(`PIVOTBY(${rp},${year},${sales},SUM,0,1,-3)`),
    t(`PIVOTBY(${reg},${year},${sales},SUM,0,1,,1.5)`),
    t(`PIVOTBY(${reg},${year},${sales},PERCENTOF,0,,,,,,1.5)`),
  ],
}

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  for (const [fn, formulas] of Object.entries(CASES)) formulas.forEach((f, i) => out.push({ id: `${fn}#${i + 1}`, group: fn, fn, formula: xl(xlpm(f)) }))
  return out
}

/** The engine's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./groupby-held-2.predicted.json', import.meta.url), 'utf8'))

/**
 * The predictions Excel for Mac did not confirm: GROUPBY takes no array of one value as the function's answer (the
 * summary is #CALC!; MAP takes it), which `INDEX(x,1)` answers there; a fractional sort order is #VALUE!.
 */
export const MISSED: readonly string[] = ['GROUPBY#2', 'GROUPBY#5', 'GROUPBY#11']
