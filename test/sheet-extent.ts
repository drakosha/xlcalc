/** Where a sheet's cells end as a result and its book give it, for the tests of `Session.getUsedRange`. */
import type { BookInput, ComputeResult, SheetExtent } from '../src/contract.js'
import { parseAddress } from '../src/model.js'
import { parseRange } from '../src/structure/area.js'
import { cellsOf } from './looks-book.js'

/** The last row and column of the cells a result lists, the cells of the sheet's looks, the spill ranges, the merged areas and the tables of the book's sheet. */
export function extentOf(r: ComputeResult, book: BookInput, sheet: string): SheetExtent {
  let rows = 0
  let columns = 0
  const reach = (row: number, col: number) => {
    rows = Math.max(rows, row + 1)
    columns = Math.max(columns, col + 1)
  }
  const area = (text: string) => {
    const x = parseRange(text)!
    reach(x.r2, x.c2)
  }
  for (const [a, c] of Object.entries(r.sheets.find((s) => s.name === sheet)!.cells)) {
    const { row, col } = parseAddress(a)
    reach(row, col)
    if (c.spill) area(c.spill)
  }
  const s = book.sheets.find((x) => x.name === sheet)!
  // a result lists no cell of the looks, which count as cells of a look alone
  for (const a of Object.keys(cellsOf(s))) area(a)
  for (const m of s.merges ?? []) area(m)
  for (const t of s.tables ?? []) area(t.range)
  return { rows, columns }
}
