/**
 * Held out, predicted by the engine's `findCells` before Excel ran it: Find from a cell (`Range.Find`'s `After`, what a
 * person's Find Next starts from), the cells from the one after it round to the first and on to it. The questions: the
 * cell to begin after inside the cells found, between them, after the last and past them all, by rows and by columns;
 * in a range, outside it, in a range of two areas; in a hidden row of a sheet without a filter (Look in Values skips
 * it), and on a sheet with a filter in a row the filter hid, in a hidden column and in a row hidden by hand, where the
 * cells come by blocks.
 */
import type { CellInput } from '../../../src/contract.js'
import { formatAddress } from '../../../src/model.js'
import { collector, type FindBook, type FindCase, type FindQuery } from './cases.js'

/** A 3×3 block with matches and three outside it, as `rules-held`'s. */
const order: Record<string, CellInput> = {
  A1: { value: 'x1' }, B1: { value: 'y' }, C1: { value: 'x2' },
  A2: { value: 'x3' }, B2: { value: 'x4' }, C2: { value: 'z' },
  A3: { value: 'y' }, B3: { value: 'x5' }, C3: { value: 'x6' },
  E1: { value: 'x7' }, A5: { value: 'x8' }, E5: { value: 'x9' },
}

/** `rules-held-3`'s: an 8×5 block, each cell `x` and its address, row 10 hidden by hand, column C hidden, a filter in column I hiding rows 3 and 6. */
function filtered(): FindBook {
  const cells: Record<string, CellInput> = { I1: { value: 'key' }, A10: { value: 'xA10' }, B10: { value: 'xB10' }, A11: { value: 'xA11' } }
  for (let r = 0; r < 8; r++) {
    for (let c = 0; c < 5; c++) cells[formatAddress(r, c)] = { value: `x${formatAddress(r, c)}` }
    if (r) cells[`I${r + 1}`] = { value: r + 1 === 3 || r + 1 === 6 ? 'drop' : 'keep' }
  }
  return { cells, rows: { 10: { hidden: true } }, columns: { C: { hidden: true } }, filter: { range: 'I1:I8', field: 1, keep: 'keep' } }
}

export const BOOKS: Record<string, FindBook> = {
  order: { cells: order },
  // row 2 hidden by hand, no filter
  hid: { cells: order, rows: { 2: { hidden: true } } },
  filtered: filtered(),
}

export function buildFindCases(): FindCase[] {
  const { out, add } = collector()
  const find = (group: string, book: string, after: string, more: Partial<FindQuery> = {}) => add(group, book, { what: 'x', lookIn: 'formulas', after, ...more })
  for (const byColumns of [false, true]) {
    const by = byColumns ? { byColumns } : {}
    // after a cell found, a cell between, the last found, past them all
    for (const after of ['B2', 'D4', 'E5', 'Z100']) find('order', 'order', after, by)
    // in a range, outside it, in a range of two areas
    find('range', 'order', 'B2', { range: 'A1:C3', ...by })
    find('range', 'order', 'A1', { range: 'B2:C3', ...by })
    find('range', 'order', 'B2', { range: 'A1:C3,E1:E5', ...by })
    find('range', 'order', 'E1', { range: 'A1:C3,E1:E5', ...by })
    // a hidden row without a filter: Look in Values skips it, Formulas does not
    for (const lookIn of ['values', 'formulas'] as const) find('hid', 'hid', 'B2', { lookIn, ...by })
    // a filtered sheet: a visible cell, a row the filter hid, a hidden column, a row hidden by hand
    for (const after of ['B2', 'B3', 'C2', 'A10']) find('filtered', 'filtered', after, by)
  }
  find('filtered', 'filtered', 'B3', { lookIn: 'values' })
  return out
}
