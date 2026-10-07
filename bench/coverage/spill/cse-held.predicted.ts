/** The answers predicted for `cse-held.ts` before Excel computed it (the engine's, by the rules then known). */
import type { Val } from '../formulajs.js'

export const PREDICTED: Record<string, Record<string, Val>> = {
  'column-longer': {
    'S!F1': { t: 'n', v: 1 },
    'S!F2': { t: 'n', v: 2 },
    'S!F3': { t: 'n', v: 3 },
    'S!F4': { t: 'e', v: 'NA' },
  },
  'outer-product': {
    'S!F6': { t: 'n', v: 10 },
    'S!G6': { t: 'n', v: 100 },
    'S!F7': { t: 'n', v: 20 },
    'S!G7': { t: 'n', v: 200 },
  },
  'sum-row': {
    'S!F9': { t: 'n', v: 15 },
    'S!G9': { t: 'n', v: 15 },
  },
  'transpose-longer': {
    'S!H1': { t: 'n', v: 10 },
    'S!I1': { t: 'n', v: 20 },
    'S!J1': { t: 'n', v: 30 },
    'S!K1': { t: 'e', v: 'NA' },
  },
  'row-times': {
    'S!H3': { t: 'n', v: 30 },
    'S!H4': { t: 'n', v: 40 },
    'S!H5': { t: 'n', v: 50 },
  },
  'single-max': {
    'S!H7': { t: 'n', v: 250 },
  },
  readers: {
    'S!I1': { t: 'n', v: 6 },
    'S!I2': { t: 'n', v: 3 },
  },
  'precedent-edit': {
    'S!J1': { t: 'n', v: 2 },
    'S!J2': { t: 'n', v: 11 },
    'S!J3': { t: 'n', v: 4 },
  },
  'part-edit': {
    'S!J5': { t: 'n', v: 1 },
    'S!J6': { t: 'n', v: 2 },
  },
}
