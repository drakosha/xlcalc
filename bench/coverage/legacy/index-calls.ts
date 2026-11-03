/**
 * INDEX with a computed row or column in formulas of the old kind (`run.ts legacy index-calls`): with a number
 * literal Excel 365 shows no `@` (the set `calls`), with a cell, `1+1`, 2.5 or TRUE it does; Excel showed a file of
 * its own with `INDEX(B1:B5,MATCH(3,A1:A5,0))` without `@`. Here each kind of call and expression as the row of a
 * column, the row and the column of a block, the row of a block alone and the column of a row; then two computed
 * places, whole columns, arrays, the area, and OFFSET, INDIRECT and CHOOSE with a computed argument. The text Excel
 * shows was predicted before Excel ran the set (`index-calls.predicted.json`).
 */
import type { Probe } from '../probes.js'
import { probe } from './params.js'
import PREDICTED from './index-calls.predicted.json' with { type: 'json' }

const M8 = 'MATCH(8,D!B1:B8,0)'

/** Rows and columns computed by calls and expressions, each 2 (or 1) over the data sheet. */
const PLACES = [
  M8, 'MATCH(2.5,D!L1:L6,1)', 'MATCH(2.5,D!L1:L6)', 'MATCH(7,{9,7,5},-1)', 'XMATCH(8,D!B1:B8)', 'ROW(D!A2)', 'ROW()', 'COLUMN(D!B1)', 'COLUMN()',
  'ROWS(D!B1:B2)', 'COLUMNS(D!B1:C1)', 'COUNT(D!B1:B2)', 'COUNTA(D!B1:B2)', 'COUNTIF(D!B1:B8,">20")', 'INT(2.5)', 'ROUND(2.4,0)', 'MAX(1,2)', 'MIN(2,3)',
  'ABS(-2)', 'LEN("ab")', 'IF(D!A4,2,1)', 'CHOOSE(1,2,3)', 'VALUE("2")', '--"2"', 'MATCH(4,D!B1:B8,0)+MATCH(4,D!B1:B8,0)', 'MATCH(15,D!B1:B8,0)-1',
  'MATCH(16,D!B1:B8,0)/2', `IFERROR(${M8},1)`, 'SUM(1,1)', 'FIND("b","ab")', 'SEARCH("b","ab")', 'MONTH("2025-02-01")', 'AREAS(D!B1)', 'FACT(2)', `(${M8})`,
  'MATCH(D!B1:B3,D!B1:B8,0)', 'INDEX(D!L1:L6,2)', 'SUMPRODUCT((D!B1:B8>20)*1)',
]

/** Where the computed place goes: the row of a column, the row and the column of a block, a block's row alone, a row's column. */
const SHAPES: Record<string, (x: string) => string> = {
  column: (x) => `INDEX(D!B1:B8,${x})`,
  row: (x) => `INDEX(D!B1:C3,${x},1)`,
  col: (x) => `INDEX(D!B1:C3,1,${x})`,
  alone: (x) => `INDEX(D!B1:C3,${x})`,
  across: (x) => `INDEX(D!B1:D1,${x})`,
}

/** Forms around them, and OFFSET, INDIRECT, CHOOSE with a computed argument. */
const FORMS = [
  `INDEX(D!B1:C3,${M8},MATCH(7,D!C1:C8,0))`, `INDEX(D!B:B,${M8})`, `INDEX(D!1:1,${M8})`, `INDEX({1;2;3},${M8})`, `INDEX({1,2;3,4},${M8},1)`, `INDEX({1,2,3},${M8})`,
  `INDEX(D!B1:C3,${M8},1,1)`, `INDEX(D!B1:C3,${M8},0)`, `INDEX(D!B1:C3,0,${M8})`, `INDEX(D!B1:C3,${M8},)`, `INDEX(D!B1:B8,${M8})+1`, `SUM(INDEX(D!B1:B8,${M8}))`,
  `SUM(D!B1:INDEX(D!B1:B8,${M8}))`, `INDEX(D!B1:B8,${M8})&""`, `INDEX(D!B1:C3,${M8},D!A12)`, `INDEX(D!B1:C3,D!A12,${M8})`, `INDEX(D!B1:C3,${M8},2.5)`,
  `OFFSET(D!B1,${M8}-1,0)`, 'OFFSET(D!B1,1,0)', `OFFSET(D!B1,0,0,${M8})`, `OFFSET(D!B1:B3,${M8}-1,0)`, `OFFSET(D!B1,${M8},0,1,1)`,
  `INDIRECT("D!B"&${M8})`, 'INDIRECT("D!B2")', `INDIRECT("D!B1:B"&${M8})`, `CHOOSE(${M8},D!B1,D!C1)`, `CHOOSE(${M8},D!B1:B3,D!C1:C3)`, `CHOOSE(${M8},1,2)`,
]

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  for (const [shape, f] of Object.entries(SHAPES)) PLACES.forEach((x, i) => out.push(probe(`${shape}#${i + 1}`, shape, 'INDEX', f(x))))
  FORMS.forEach((f, i) => out.push(probe(`form#${i + 1}`, 'form', 'INDEX', f)))
  return out
}

/** The text Excel 365 was predicted to show for each case (canonical, without `=`), committed before Excel ran. */
export const PREDICTED_SHOWN: Readonly<Record<string, string>> = PREDICTED

/**
 * The predictions Excel did not confirm: of the calls predicted to leave one cell only MATCH does (XMATCH, ROW, COLUMN,
 * ROWS, COLUMNS get `@`), and OFFSET over one cell with no size or the size 1 is one cell.
 */
export const MISSED: readonly string[] = [...Object.keys(SHAPES).flatMap((s) => [5, 6, 7, 8, 9, 10, 11].map((i) => `${s}#${i}`)), 'form#18', 'form#19', 'form#22']
