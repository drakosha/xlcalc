/**
 * Held out: a spill reference typed in INDIRECT's text. The sets `indirect` and `indirect-held` sent it to Excel as
 * `INDIRECT("_xlfn.ANCHORARRAY(R!E1)")`, the harness writing the file form inside the quotes too (`anchorArray` in
 * `book.ts`, fixed since), which Excel answered `#REF!`. Answers: `test/excel/cases/indirect-held-3.json`;
 * `PREDICTED`, the engine's answers, was committed before Excel ran the set.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe } from '../probes.js'
import { collector } from './arrays/common.js'
import { SHEETS } from './indirect.js'

export { SHEETS }

const FORMULAS = ['SUM(INDIRECT("R!E1#"))', 'ROWS(INDIRECT("R!E1#"))', 'INDEX(INDIRECT("R!E1#"),2)', 'SUM(INDIRECT("R!A1#"))', 'SUM(INDIRECT("R!E2#"))', 'SUM(OFFSET(INDIRECT("R!E1#"),1,0))', 'COUNT(INDIRECT("R!$E$1#"))']

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  add('spill', 'INDIRECT', ...FORMULAS)
  return out
}

/** The engine's answers by case id, committed before Excel ran the set (`predict.ts`). */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./indirect-held-3.predicted.json', import.meta.url), 'utf8'))
