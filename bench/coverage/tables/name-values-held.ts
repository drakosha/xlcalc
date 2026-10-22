/**
 * Held out: the values of cells that use names after the structure operations of the struct set `names-held` (the
 * same books and operations, whose formulas and names that set reads), predicted by the engine before Excel took them
 * (`name-values-held.predicted.json`): a sheet's own name referenced as `S!loc` when the sheet is renamed or deleted,
 * and names that mix `$` and offsets beside cell formulas reading the same cells when rows and columns are inserted and deleted.
 */
import { BOOKS, OPS } from '../struct/names-held.js'
import type { RangeRef } from '../../../src/contract.js'
import type { TableCase } from './cases.js'

/** Where the cells are after the operation: every sheet the book keeps, over the cells the operations move. */
function readOf(book: string, sheets: string[]): RangeRef[] {
  return sheets.map((sheet) => ({ sheet, range: book === 'qualified' ? 'A1:B4' : 'A1:H22' }))
}

export function buildTableCases(): TableCase[] {
  return Object.entries(OPS).flatMap(([book, ops]) =>
    ops.map(([id, op]): TableCase => {
      const b = BOOKS[book]()
      const sheets = b.sheets.map((s) => s.name).flatMap((s) => (op.op === 'deleteSheet' && op.sheet === s ? [] : op.op === 'renameSheet' && op.sheet === s ? [op.name] : [s]))
      return { id: `${book}-${id}`, about: `${book}: ${JSON.stringify(op)}`, book: b, steps: [{ op }], read: readOf(book, sheets) }
    }),
  )
}
