/** The answers predicted for `legacy-held.ts` before Excel computed it (the engine's, by the rules then known). */
import type { Val } from '../formulajs.js'

export const PREDICTED: Record<string, Record<string, Val>> = {
  positions: {
    'T!B4': { t: 'n', v: 160 },
    'S!H6': { t: 's', v: 'big' },
    'S!H2': { t: 'n', v: 400 },
    'S!H3': { t: 'n', v: 30 },
    'S!H4': { t: 'n', v: 80 },
    'S!H5': { t: 'n', v: 2 },
    'S!H7': { t: 'n', v: 12 },
    'S!H8': { t: 'n', v: 16 },
    'S!H9': { t: 'n', v: 90 },
    'S!H11': { t: 'e', v: 'VALUE' },
    'T!E2': { t: 'n', v: 55 },
    'S!H12': { t: 'e', v: 'VALUE' },
    'S!H10': { t: 'n', v: 30 },
  },
}
