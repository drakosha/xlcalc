/**
 * Held out, predicted by the engine's `findCells` before Excel ran it, after `rules-held-4` showed that on a filtered
 * sheet Find from a cell of a hidden row or column (or from the first block's last cell) goes on with the second block,
 * as if the cell stood in the range's first block. The questions: Find from a cell of each later block, visible, and from
 * cells of hidden rows and columns further on, by rows and by columns, over the sheet and in ranges that begin past the
 * first block.
 */
import { collector, type FindBook, type FindCase, type FindQuery } from './cases.js'
import { BOOKS as HELD_4 } from './rules-held-4.js'

export const BOOKS: Record<string, FindBook> = { filtered: HELD_4.filtered }

export function buildFindCases(): FindCase[] {
  const { out, add } = collector()
  const find = (after: string, more: Partial<FindQuery> = {}) => add('filtered', 'filtered', { what: 'x', lookIn: 'formulas', after, ...more })
  for (const byColumns of [false, true]) {
    const by = byColumns ? { byColumns } : {}
    // visible cells of later blocks, the last block's last cell
    for (const after of ['D1', 'A4', 'D4', 'E5', 'A11']) find(after, by)
    // hidden rows and a hidden column further on
    for (const after of ['D3', 'A6', 'C5', 'E10']) find(after, by)
    // ranges that begin past the first block
    find('D4', { range: 'B2:E8', ...by })
    find('B3', { range: 'B2:E8', ...by })
    find('C7', { range: 'A4:E8', ...by })
  }
  return out
}
