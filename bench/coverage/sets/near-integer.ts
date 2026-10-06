/**
 * Numbers near a whole number, as dynamic array formulas (Excel 365): where Excel takes a number a little
 * below (above) an integer as that integer. The sheet NI holds exact values near 1, 2, 3, 10, 10^6, 10^15, a
 * date serial, negative ones and 0.3 (`near-integer/values.ts`); the cases read them in whole-number arguments,
 * INT and its kin, MOD, comparisons, criteria and lookups, as text and as dates, first on a ladder of relative
 * distances (`ladder.ts`), then around the boundaries it found (`fine.ts`, `tolerance.ts`). Answers: `test/excel/cases/near-integer.json`.
 */
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'
import { fine } from './near-integer/fine.js'
import { ladder } from './near-integer/ladder.js'
import { tolerance } from './near-integer/tolerance.js'
import { values } from './near-integer/values.js'

export const DYNAMIC = true

const { sheet, g } = values()

export const SHEETS: SheetCells = { NI: sheet }

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  ladder(add, g)
  fine(add, g)
  tolerance(add, g)
  return out
}
