/** The answers predicted for `names-held.ts` before Excel computed it (the engine's, by the rules then known). */
import type { Val } from '../formulajs.js'

export const PREDICTED: Record<string, Record<string, Val>> = {
  'relative-other-sheet': {
    'T!C4': { t: 'n', v: 40 },
  },
  'union-area': {
    'S!G1': { t: 'n', v: 4 },
    'S!G2': { t: 'n', v: 2 },
    'S!G3': { t: 'n', v: 33 },
    'S!G4': { t: 'n', v: 8 },
  },
  'array-arith': {
    'S!G1': { t: 'n', v: 1 },
    'S!H1': { t: 'n', v: 2 },
    'S!I1': { t: 'n', v: 3 },
    'S!G3': { t: 'n', v: 1400 },
  },
  'index-criteria': {
    'S!G1': { t: 'n', v: 120 },
    'S!G2': { t: 'n', v: 3 },
    'S!G3': { t: 'n', v: 50 },
  },
  'qualified-local': {
    'S!G1': { t: 'n', v: 10 },
    'S!G2': { t: 'n', v: 1 },
  },
  'let-in-name': {
    'S!G1': { t: 'n', v: 14 },
    'S!H1': { t: 'n', v: 1 },
    'S!H2': { t: 'n', v: 4 },
    'S!H3': { t: 'n', v: 9 },
  },
  'relative-chain': {
    'S!C3': { t: 'n', v: 300 },
  },
  'spill-ref': {
    'S!G1': { t: 'n', v: 6 },
    'S!G2': { t: 'n', v: 3 },
  },
  'at-name': {
    'S!G3': { t: 'n', v: 3 },
    'S!G7': { t: 'e', v: 'VALUE' },
  },
  'formula-twice': {
    'S!G1': { t: 'n', v: 10 },
  },
}
