/** The answers predicted for `at-held.ts` before Excel computed it (the engine's, by the rules then known). */
import type { Val } from '../formulajs.js'

export const PREDICTED: Record<string, Record<string, Val>> = {
  mixed: {
    'S!H1': { t: 'n', v: 7 },
    'S!I2': { t: 'n', v: 200 },
    'S!H3': { t: 'n', v: 330 },
    'S!I3': { t: 's', v: 'c' },
    'S!H4': { t: 'n', v: 5 },
    'S!I5': { t: 'n', v: 1 },
    'S!H9': { t: 'n', v: 1 },
  },
}
