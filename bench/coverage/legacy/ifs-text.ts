/**
 * IFS with a condition of text in formulas of the old kind (`run.ts legacy ifs-text`): Excel 365 counts its #VALUE!
 * as one cell in ROWS (the case set `known-held` measures the dynamic kind); here whether the old kind does too, and
 * the controls of an error given as a value. `PREDICTED`, the engine's answers to the converted formulas, was
 * committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { Probe } from '../probes.js'
import { probe } from './params.js'
import { n, e } from '../vals.js'

/** Formula and the answer predicted before Excel ran. */
const CASES: [string, Val][] = [
  ['IFS("x",1,TRUE,2)', e('VALUE')],
  ['ROWS(IFS("x",1,TRUE,2))', n(1)],
  ['COLUMNS(IFS(FALSE,1,"x",2))', n(1)],
  ['TYPE(IFS("x",1))', n(16)],
  ['ROWS(IFS(D!A3,1))', n(1)],
  ['SUM(IFS("x",1))', e('VALUE')],
  ['ROWS(VSTACK(1,IFS("x",1)))', n(2)],
  ['ROWS(IFS(TRUE,1/0))', e('DIV0')],
]

export function buildProbes(): Probe[] {
  return CASES.map(([f], i) => probe(`IFS#${i + 1}`, 'ifs', 'IFS', f))
}

/** The predicted answers by case id. */
export const PREDICTED: Readonly<Record<string, Val>> = Object.fromEntries(buildProbes().map((p, i) => [p.id, CASES[i][1]]))

/** The predictions Excel did not confirm: Excel 365 shows `TYPE(IFS("x",1))` with no `@`, so TYPE takes IFS's array whole (64). */
export const MISSED: readonly string[] = ['IFS#4']
