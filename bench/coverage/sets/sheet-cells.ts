/**
 * Sheets whose names read as cell addresses (S1, S2, S3): the unquoted `S1:S3!A1`, which Excel does not read as a
 * reference across sheets (the set `areas`), against the quoted `'S1:S3'!A1`, on the case sheet and on the sheets
 * S1 and S3 themselves, where `S1` is a cell of the formula's own sheet. Answers: `test/excel/cases/sheet-cells.json`;
 * `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'

/** S1, S2, S3 in this order before D and C; formulas with the unquoted names on S1 and S3. */
export const SHEETS: SheetCells = {
  S1: { A1: 1, A2: 2, C5: { formula: 'SUM(S1:S3!A1)' } },
  S2: { A1: 4, A2: 5 },
  S3: { A1: 7, A2: 8, B1: 70, C5: { formula: 'SUM(S1:S3!A1)' }, C6: { formula: 'SUM(S1:S3!A1:A2)' } },
}

const n = (v: number): Val => ({ t: 'n', v })
const b = (v: boolean): Val => ({ t: 'b', v })
const e = (v: Extract<Val, { t: 'e' }>['v']): Val => ({ t: 'e', v })

/** Formula and the answer predicted before Excel ran (undefined: not predicted). */
const CASES: [string, Val?][] = [
  ['S1:S3!A1', e('VALUE')],
  ['SUM(S1:S3!A1)', e('VALUE')],
  ['ISREF(S1:S3!A1)', b(false)],
  ['SUM(S1:S3!A1:A2)', e('VALUE')],
  ["'S1:S3'!A1", e('REF')],
  ["SUM('S1:S3'!A1)", n(12)],
  ["SUM('S1:S3'!A1:A2)", n(27)],
  ["'S1'!A1", n(1)],
  ["'S1'!C5", e('VALUE')],
  ["'S3'!C5", n(77)],
  ["'S3'!C6", n(85)],
  ['S1!A1'],
  ['SUM(S2:S3!A1)'],
]

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [f] of CASES) add('sheets', 'sheets', f)
  return out
}

/** The predicted answers by case id. */
export const PREDICTED: Readonly<Record<string, Val>> = Object.fromEntries(
  buildProbes().flatMap((p, i) => (CASES[i][1] ? [[p.id, CASES[i][1]]] : [])),
)
