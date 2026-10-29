/**
 * GROUPBY and PIVOTBY over a small table (sheet G: region, product, year, sales, quantity, a flag; a header row, a region
 * in another case, a blank region, text among the sales): headers found, shown, made up; totals and subtotals at the
 * bottom and the top; sorting by fields and by values; a filter; several value columns and functions, labelled; a
 * LAMBDA, PERCENTOF and its totals; the table relationship; wrong arguments. A summary is read through ARRAYTOTEXT
 * (strict), so its size and every cell are one case, an empty cell told from empty text. Answers:
 * `test/excel/cases/groupby.json`; `PREDICTED`, the engine's answers (`sets/predict.ts groupby`), was committed before
 * Excel ran the set.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { xl } from './arrays/common.js'
import { xlpm } from './let/xlpm.js'

const ROWS: (string | number | boolean | null)[][] = [
  ['Region', 'Product', 'Year', 'Sales', 'Qty', 'Keep'],
  ['East', 'A', 2024, 10, 1, true],
  ['West', 'B', 2024, 20, 2, false],
  ['East', 'B', 2025, 5, 3, true],
  ['west', 'A', 2025, 7, 4, true],
  ['North', 'A', 2024, 3, 5, true],
  ['East', 'A', 2025, 8, 6, false],
  [null, 'B', 2024, 4, 7, true],
  ['North', 'C', 2025, 'x', 8, true],
  ['West', 'C', 2024, 12, 9, true],
  ['East', 'C', 2025, 6, 10, true],
]

export const SHEETS: SheetCells = {
  G: Object.fromEntries(ROWS.flatMap((row, r) => row.flatMap((v, c) => (v === null ? [] : [[`${'ABCDEF'[c]}${r + 1}`, v]])))),
}

const t = (f: string) => `ARRAYTOTEXT(${f},1)`
const [reg, regH, rp, rpH, year, sales, salesH, sq, sqH, keep] = ['G!A2:A11', 'G!A1:A11', 'G!A2:B11', 'G!A1:B11', 'G!C2:C11', 'G!D2:D11', 'G!D1:D11', 'G!D2:E11', 'G!D1:E11', 'G!F2:F11']

const CASES: Record<string, string[]> = {
  GROUPBY: [
    t(`GROUPBY(${reg},${sales},SUM)`),
    t(`GROUPBY(${regH},${salesH},SUM)`),
    t(`GROUPBY(${regH},${salesH},SUM,3)`),
    t(`GROUPBY(${reg},${sales},SUM,2)`),
    t(`GROUPBY(${rpH},${salesH},SUM)`),
    t(`GROUPBY(${rp},${sales},SUM,0,1)`),
    t(`GROUPBY(${rp},${sales},SUM,0,2)`),
    t(`GROUPBY(${rp},${sales},SUM,0,-2)`),
    t(`GROUPBY(${rp},${sales},SUM,0,0)`),
    t(`GROUPBY(${reg},${sales},SUM,0,2)`),
    t(`GROUPBY(${reg},${sales},SUM,,,-2)`),
    t(`GROUPBY(${reg},${sales},SUM,,,-1)`),
    t(`GROUPBY(${reg},${sales},SUM,,,,${keep})`),
    t(`GROUPBY(${reg},${sq},SUM)`),
    t(`GROUPBY(${regH},${sqH},SUM)`),
    t(`GROUPBY(${reg},${sales},HSTACK(SUM,COUNT))`),
    t(`GROUPBY(${reg},${sales},VSTACK(SUM,COUNT))`),
    t(`GROUPBY(${reg},${sales},LAMBDA(x,MAX(x)))`),
    t(`GROUPBY(${reg},${sales},PERCENTOF)`),
    t(`GROUPBY(${reg},${sales},COUNTA)`),
    t(`GROUPBY(${rp},${sales},SUM,0,2,,,1)`),
    t(`GROUPBY(${rp},${sales},SUM,0,1,,,1)`),
    t(`GROUPBY(${reg},${sales},1)`),
    t(`GROUPBY(G!A2:A10,${sales},SUM)`),
    t(`GROUPBY(${reg},${sales},SUM,,,,${keep}=99)`),
    t(`GROUPBY(${year},${sales},SUM)`),
    t(`GROUPBY(${reg},${sales},LAMBDA(x,x))`),
    t(`GROUPBY(${rp},${sales},SUM,0,1,{2,-3})`),
    t(`GROUPBY(${reg},${sales},SUM,3)`),
    t(`GROUPBY(${reg},${sales},VSTACK(HSTACK(SUM,COUNT),{"s","c"}))`),
    `ROWS(GROUPBY(${reg},${sales},SUM))`,
    t(`GROUPBY(${reg},${sales},SUM,4)`),
    t(`GROUPBY(${reg},${sales},AVERAGE)`),
    t('GROUPBY({"a";"b";"a"},{1;2;3},SUM)'),
    t(`GROUPBY(${rp},${sales},SUM)`),
    t(`GROUPBY(${rp},${sales},SUM,,,,,1)`),
  ],
  PIVOTBY: [
    t(`PIVOTBY(${reg},${year},${sales},SUM)`),
    t(`PIVOTBY(${regH},G!C1:C11,${salesH},SUM)`),
    t(`PIVOTBY(${regH},G!C1:C11,${salesH},SUM,3)`),
    t(`PIVOTBY(${reg},${year},${sales},SUM,0,0,,0)`),
    t(`PIVOTBY(${reg},${year},${sales},PERCENTOF)`),
    t(`PIVOTBY(${reg},${year},${sales},PERCENTOF,,,,,,,1)`),
    t(`PIVOTBY(${reg},${year},${sales},PERCENTOF,,,,,,,2)`),
    t(`PIVOTBY(${reg},G!B2:C11,${sales},SUM)`),
    t(`PIVOTBY(${reg},${year},${sq},SUM)`),
    t(`PIVOTBY(${reg},${year},${sales},HSTACK(SUM,COUNT))`),
    t(`PIVOTBY(${reg},${year},${sales},SUM,,,-2)`),
    t(`PIVOTBY(${reg},${year},${sales},SUM,,,,,-1)`),
    t(`PIVOTBY(${rp},${year},${sales},SUM)`),
    t(`PIVOTBY(${reg},${year},${sales},SUM,,,,,,${keep})`),
    t(`PIVOTBY(${reg},${year},${sales},COUNT)`),
  ],
}

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  for (const [fn, formulas] of Object.entries(CASES)) formulas.forEach((f, i) => out.push({ id: `${fn}#${i + 1}`, group: fn, fn, formula: xl(xlpm(f)) }))
  return out
}

/** The engine's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./groupby.predicted.json', import.meta.url), 'utf8'))
