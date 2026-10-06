/**
 * Open cases of the random check (the `fuzz` set): differences from Excel whose cause lies in engine files
 * outside the case work (the compiler, the runtime, the array layer), each with one line why and the engine's
 * answer pinned, as KNOWN does. The lead closes them; an entry whose case matches Excel fails the tests.
 */
import type { Val } from '../bench/coverage/formulajs.js'

const LEFT_FIRST = 'an operator converts its left operand before the right one\'s error counts: `"abc"-NA()` is #VALUE! in Excel, #N/A here (src/runtime.ts, arith)'

export const OPEN: Readonly<Record<string, { why: string; engine: Val }>> = {
  'fuzz:1#253': { why: LEFT_FIRST, engine: { t: 'e', v: 'NA' } },
  'fuzz:1#878': { why: LEFT_FIRST, engine: { t: 'e', v: 'NA' } },
  'fuzz:2#1430': { why: LEFT_FIRST, engine: { t: 'e', v: 'NA' } },
  'fuzz:2#1640': { why: LEFT_FIRST, engine: { t: 'e', v: 'CALC' } },
  'fuzz:2#285': {
    why: 'an operator over a 1×1 array answers a 1×1 array in Excel (`TYPE({1}+0)` is 64, so `DEVSQ({1}>=TRUE)` skips the boolean: #NUM!); here it answers a value (src/runtime.ts, L1 and L2 take a 1×1 array as its value)',
    engine: { t: 'n', v: 0 },
  },
  'fuzz:1#1790': {
    why: 'VSTACK and HSTACK answer an error given as a value (`INDEX(VSTACK(1,1/0,2),1)` is #DIV/0!) but stack a one-cell reference to an error (`HSTACK(D!B1:B2,D!E3)`); telling them apart needs `p` arguments passed as references (src/compile.ts)',
    engine: { t: 'n', v: 6 },
  },
}
