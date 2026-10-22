/**
 * Open cases of the random check (the `fuzz` set) and of case sets: differences from Excel whose cause lies in engine files
 * outside the case work (the compiler, the runtime, the array layer), each with one line why and the engine's
 * answer pinned, as KNOWN does. The lead closes them; an entry whose case matches Excel fails the tests.
 */
import type { Val } from '../bench/coverage/formulajs.js'

/** SUMIF inside its own resized sum range: Excel finds the circular reference and leaves the formulas reading the cell uncomputed (empty), so its value is not read; the engine holds the cycle at 0. */
const SUMIF_OWN = "a circular reference through SUMIF's resized sum range: Excel leaves the formulas that read the cell uncomputed, its value unread; the engine holds the cycle at 0"

export const OPEN: Readonly<Record<string, { why: string; engine: Val }>> = {
  'sumif-own-held:SUMIF#1': { why: SUMIF_OWN, engine: { t: 'n', v: 0 } },
  'sumif-own-held:SUMIF#2': { why: SUMIF_OWN, engine: { t: 'n', v: 0 } },
  'sumif-own-held:SUMIF#3': { why: SUMIF_OWN, engine: { t: 'n', v: 0 } },
}
