/**
 * A book's looks (`SheetInput.looks`) as entries of a look alone, as a book held them before the looks: for the tests
 * that read a book's cells as entries, and the book a book with looks is compared with.
 */
import type { BookInput, CellInput, SheetInput } from '../src/contract.js'
import { readRanges } from '../src/looks/ranges.js'
import { formatAddress } from '../src/model.js'

/** The sheet with each cell of a look of its looks and without an entry as an entry of that look, and no looks. */
export function expandSheet(sheet: SheetInput): SheetInput {
  if (!sheet.looks) return sheet
  const { looks, ...rest } = sheet
  const cells: Record<string, CellInput> = { ...sheet.cells }
  const layer = readRanges(looks)
  for (const p of layer.pieces())
    for (let r = p.r1; r <= p.r2; r++)
      for (let c = p.c1; c <= p.c2; c++) {
        const a = formatAddress(r, c)
        if (!cells[a]) cells[a] = { ...layer.looks.list[p.look] }
      }
  return { ...rest, cells }
}

/** The book with every sheet's looks as entries (`expandSheet`). */
export const expandBook = <B extends Pick<BookInput, 'sheets'>>(book: B): B => ({ ...book, sheets: book.sheets.map(expandSheet) })

/** A sheet's cells with its looks as entries (`expandSheet`). */
export const cellsOf = (sheet: SheetInput): Record<string, CellInput> => expandSheet(sheet).cells
