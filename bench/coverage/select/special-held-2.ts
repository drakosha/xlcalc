/**
 * Held out after special-held: visible cells over one cell (the whole sheet's) and over an entire column, with hidden
 * rows and columns. Predicted before Excel ran it (`special-held-2.predicted.json`).
 */
import type { SelectCase } from './cases.js'
import { BOOKS as HELD } from './special-held.js'

export const BOOKS = { lines: HELD.lines }

export function buildSelectCases(): SelectCase[] {
  return [
    { id: 'visible-one-cell', book: 'lines', range: 'A1', special: { type: 'visible' } },
    { id: 'visible-column', book: 'lines', range: 'D:D', special: { type: 'visible' } },
  ]
}
