/**
 * SUMIF beside the cells its sum range takes in when resized to the criteria range, held out: a SUMIF inside its own
 * resized range, a formula there that reads the SUMIF, and one there that does not. Each formula stands on the set's
 * sheets and a case reads its cell. `PREDICTED`, the engine's answers, was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { SheetCells } from '../probes.js'
import { heldCases } from './arrays/common.js'
import { n } from '../vals.js'

const sumif = { formula: 'SUMIF(C1:C3,"",B1)', dynamic: true as const }

/** Ua: B2 sums B1:B3, itself among them; Ub: B3 reads B2 as well; Uc: D1 sums B1:B3, B2 a formula that reads nothing. */
export const SHEETS: SheetCells = {
  Ua: { B1: 1, B3: 2, B2: sumif },
  Ub: { B1: 1, B2: sumif, B3: { formula: 'B2*2', dynamic: true } },
  Uc: { B1: 1, B2: { formula: '2+3', dynamic: true }, B3: 2, D1: { ...sumif } },
}

/** The cell read and the answer predicted before Excel ran: 0 where the engine finds a cycle and holds the cell at 0. */
const CASES: [string, Val][] = [
  ['Ua!B2', n(0)],
  ['Ub!B2', n(0)],
  ['Ub!B3', n(0)],
  ['Uc!D1', n(8)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES.map(([f, v]) => ['SUMIF', f, v]))

/** The predictions Excel did not confirm: it left the formulas reading a cell in the circular reference uncomputed (`OPEN` in `test/excel-open.ts`). */
export const MISSED = ['SUMIF#1', 'SUMIF#2', 'SUMIF#3']
