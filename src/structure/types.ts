/**
 * Structure operations on a book (`restructure`): inserting and deleting rows and columns, moving and
 * copying ranges, renaming and deleting sheets, with every reference in formulas and names adjusted as
 * Excel adjusts it. Not part of `./contract` yet.
 */
import type { BookInput, CellRef } from '../contract.js'

/** A rectangle on a sheet: `B2`, `B2:D9`, whole columns `B:D` or whole rows `2:9` (no `$`, no sheet). */
export interface RangeRef {
  sheet: string
  range: string
}

/**
 * One structure operation. Rows are 1-based numbers and columns are letters, as in A1 addresses;
 * `count` is 1 by default.
 */
export type StructureOp =
  /** Inserts `count` empty rows before `row`: rows from `row` on move down. */
  | { op: 'insertRows'; sheet: string; row: number; count?: number }
  /** Deletes rows `row` … `row + count - 1`: the rows below move up. */
  | { op: 'deleteRows'; sheet: string; row: number; count?: number }
  /** Inserts `count` empty columns before `column`: columns from `column` on move right. */
  | { op: 'insertColumns'; sheet: string; column: string; count?: number }
  /** Deletes `count` columns from `column` on: the columns to the right move left. */
  | { op: 'deleteColumns'; sheet: string; column: string; count?: number }
  /** Cut and paste: the range moves so that its top-left cell lands on `to`, replacing what was there. */
  | { op: 'moveRange'; from: RangeRef; to: CellRef }
  /**
   * Copy and paste: the range is pasted at `to`; a target that is a whole multiple of the source's size
   * is tiled with copies (fill down and fill right are copies of the first row or column).
   */
  | { op: 'copyRange'; from: RangeRef; to: RangeRef }
  | { op: 'renameSheet'; sheet: string; name: string }
  | { op: 'deleteSheet'; sheet: string }

/** Why Excel would refuse an operation. */
export type StructureRefusal =
  /** Cells that are not empty would leave the sheet (past XFD or row 1,048,576), or a paste does not fit. */
  | 'off_sheet'
  /** A copy target that is not a whole multiple of the source's size. */
  | 'paste_size'
  /** A new sheet name that is empty or taken (sheet names compare case-insensitively). */
  | 'sheet_name'
  /** The only sheet of a book cannot be deleted. */
  | 'last_sheet'

export type RestructureResult =
  | {
      ok: true
      /** The new book; the input book is not changed, unchanged sheets and cells are shared. */
      book: BookInput
      /** Where a cell of the old book is now; null if it was deleted or overwritten. */
      mapCell(ref: CellRef): CellRef | null
    }
  | { ok: false; code: StructureRefusal; message: string }
