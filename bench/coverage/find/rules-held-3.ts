/**
 * Held out, predicted by the engine's `findCells` before Excel ran it, after `rules-held-2` showed that Look in Formulas
 * finds the cells of hidden rows and columns on a sheet without a filter, and takes them in plain order, where in
 * `rules-held` (a sheet with a filter) it skipped them and took the cells by blocks. The question: whether a sheet whose
 * filter has criteria (Excel's filter mode) is what makes Find skip hidden cells in Look in Formulas too and take the
 * visible cells by blocks, rows of blocks first, by rows and by columns, over the sheet and in ranges, a range away from
 * the filter's rows too; and whether a filter whose criteria hide no row does the same.
 */
import type { CellInput } from '../../../src/contract.js'
import { formatAddress } from '../../../src/model.js'
import { collector, type FindBook, type FindCase } from './cases.js'

/** An 8×5 block, each cell `x` and its address, a row 10 hidden by hand past the filter's rows and a row 11 under it, the filter's key in column I. */
const book = (drop: number[]): FindBook => {
  const cells: Record<string, CellInput> = { I1: { value: 'key' }, A10: { value: 'xA10' }, B10: { value: 'xB10' }, A11: { value: 'xA11' } }
  for (let r = 0; r < 8; r++) {
    for (let c = 0; c < 5; c++) cells[formatAddress(r, c)] = { value: `x${formatAddress(r, c)}` }
    if (r) cells[`I${r + 1}`] = { value: drop.includes(r + 1) ? 'drop' : 'keep' }
  }
  return { cells, rows: { 10: { hidden: true } }, columns: { C: { hidden: true } }, filter: { range: 'I1:I8', field: 1, keep: 'keep' } }
}

export const BOOKS: Record<string, FindBook> = {
  // the filter hides rows 3 and 6
  filtered: book([3, 6]),
  // the filter's criteria hide no row
  quiet: book([]),
}

export function buildFindCases(): FindCase[] {
  const { out, add } = collector()
  for (const b of Object.keys(BOOKS))
    for (const range of [undefined, 'A1:E5', 'B2:E11', 'A9:E11'])
      for (const byColumns of [false, true])
        for (const lookIn of ['formulas', 'values'] as const) add(b, b, { what: 'x', lookIn, ...(range ? { range } : {}), ...(byColumns ? { byColumns } : {}) })
  return out
}
