/**
 * Open cases of the random check (the `fuzz` set) and of case sets: differences from Excel whose cause lies in engine files
 * outside the case work (the compiler, the runtime, the array layer), each with one line why and the engine's
 * answer pinned, as KNOWN does. The lead closes them; an entry whose case matches Excel fails the tests.
 */
import type { Val } from '../bench/coverage/formulajs.js'

/** SUMIF inside its own resized sum range: Excel finds the circular reference and leaves the formulas reading the cell uncomputed (empty), so its value is not read; the engine holds the cycle at 0. */
const SUMIF_OWN = "a circular reference through SUMIF's resized sum range: Excel leaves the formulas that read the cell uncomputed, its value unread; the engine holds the cycle at 0"

/** Excel for Mac 16.114 answered #NUM! in the first run of the set `indirect`, a run beside other workbooks' calculations; the same formulas in the book of `indirect-held` gave the engine's answer. */
const MAC_NUM = 'Excel for Mac 16.114 answered #NUM! in the first run of indirect, not in the book of indirect-held (the same formulas, again#1 to again#6): not repeated'

export const OPEN: Readonly<Record<string, { why: string; engine: Val }>> = {
  'indirect:a1#26': { why: MAC_NUM, engine: { t: 'e', v: 'DIV0' } },
  'indirect:offset#37': { why: MAC_NUM, engine: { t: 'n', v: 30 } },
  'indirect:arrays#1': { why: MAC_NUM, engine: { t: 'n', v: 30 } },
  'indirect:arrays#2': { why: MAC_NUM, engine: { t: 'n', v: 40 } },
  'indirect:arrays#3': { why: MAC_NUM, engine: { t: 'n', v: 10 } },
  'indirect:arrays#4': { why: MAC_NUM, engine: { t: 'n', v: 40 } },
  'indirect:arrays#5': { why: MAC_NUM, engine: { t: 'n', v: 9 } },
  'indirect:cell#26': { why: MAC_NUM, engine: { t: 'n', v: 5 } },
  'indirect:formulatext#4': { why: MAC_NUM, engine: { t: 'e', v: 'NA' } },
  'sumif-own-held:SUMIF#1': { why: SUMIF_OWN, engine: { t: 'n', v: 0 } },
  'sumif-own-held:SUMIF#2': { why: SUMIF_OWN, engine: { t: 'n', v: 0 } },
  'sumif-own-held:SUMIF#3': { why: SUMIF_OWN, engine: { t: 'n', v: 0 } },
}
