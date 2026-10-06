/**
 * Open cases of the random check (the `fuzz` set): differences from Excel whose cause lies in engine files
 * outside the case work (the compiler, the runtime, the array layer), each with one line why and the engine's
 * answer pinned, as KNOWN does. The lead closes them; an entry whose case matches Excel fails the tests.
 */
import type { Val } from '../bench/coverage/formulajs.js'

export const OPEN: Readonly<Record<string, { why: string; engine: Val }>> = {
  'fuzz:1#1790': {
    why: 'VSTACK and HSTACK answer an error given as a value (`INDEX(VSTACK(1,1/0,2),1)` is #DIV/0!) but stack a one-cell reference to an error (`HSTACK(D!B1:B2,D!E3)`); telling them apart needs `p` arguments passed as references (src/compile.ts)',
    engine: { t: 'n', v: 6 },
  },
}
