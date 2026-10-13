/**
 * Held out, predicted before Excel ran it by `getEnd` of the day, after `rules-held` confirmed every prediction: `End`
 * from a cell in a hidden row or column (one with a value and one without), rows of height 0 and columns of width 0
 * that are not marked hidden, and a merged area whose top-left cell is in a hidden row.
 */
import { collector, named, type EndBook, type EndCase } from './cases.js'

export const BOOKS: Record<string, EndBook> = {
  // rows 5 and 6 and column C hidden, A5 and C1 with a value
  start: {
    cells: named('A1', 'A2', 'A3', 'A4', 'A5', 'A7', 'A8', 'A9', 'A11', 'B1', 'C1', 'D1', 'E1', 'G1'),
    rows: { '5:6': { hidden: true } },
    columns: { C: { hidden: true } },
  },
  // row 3 of height 0 and column B of width 0, neither marked hidden
  sizes: {
    cells: named('A1', 'A2', 'A4', 'A5', 'A7', 'C1', 'D1', 'F1'),
    rows: { 3: { height: 0 } },
    columns: { B: { width: 0 } },
  },
  // B8:B10 merged with its value in B8, row 8 hidden
  merged: {
    cells: named('B6', 'B8', 'B12'),
    merges: ['B8:B10'],
    rows: { 8: { hidden: true } },
  },
}

export function buildEndCases(): EndCase[] {
  const { out, list } = collector()
  list('start', 'start', 'A5:d A6:d A5:u A6:u C1:r C1:l C2:r C2:l A6:r A5:r')
  list('sizes', 'sizes', 'A1:d A5:u A1:r F1:l A3:d')
  list('merged', 'merged', 'B6:d B12:u B9:u B9:d')
  return out
}
