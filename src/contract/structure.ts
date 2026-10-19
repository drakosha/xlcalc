/** Structure operations (`editStructure`) and their answers, and a sheet's outline groups and commands over its rows and columns (`outlineGroups`, `outlineEdit`). */
import type { CellRef, LineSpan, RangeRef, Refusal } from './values.js'
import type { Book, ColumnInput, RowInput, TableInput } from './book.js'
import type { EditResult } from './edits.js'

/**
 * A group of rows or columns as Excel draws it (`outlineGroups`): the longest run of lines `first` to `last` (columns by
 * number) at outline level `level` or deeper, so a group of level 2 lies inside one of level 1.
 */
export interface OutlineGroup extends LineSpan {
  level: number
  /**
   * The summary line, which holds the group's button: the line past it, below or right (before it with summaries above or
   * left, `OutlineSettings`); null where that is past the sheet's edge, where Excel draws no button. Groups that end
   * together share one.
   */
  summary: number | null
  /** Collapsed, the button showing `+`: its summary line marked as Excel keeps the mark (`RowInput.collapsed`); without a summary line, all its lines hidden. */
  collapsed: boolean
}

/** A sheet's groups (`outlineGroups`), each axis by first line, the outer group before the inner where two begin together. */
export interface SheetOutline {
  rows: OutlineGroup[]
  columns: OutlineGroup[]
}

/** An outline operation of `StructureOp` without its sheet, for a sheet alone (`outlineEdit`): `{ op: 'group', rows: '5:7' }`. */
export type OutlineEditOp = WithoutSheet<Extract<StructureOp, { op: 'group' | 'ungroup' | 'showDetail' | 'hideDetail' | 'showLevels' }>>

/** Each operation of a union without its `sheet`. */
type WithoutSheet<T> = T extends unknown ? Omit<T, 'sheet'> : never

/**
 * A sheet's rows and columns after an outline command (`outlineEdit`), each axis whole as `SheetInput` holds it and left
 * out when it has no lines; or why Excel refuses the command, as `editStructure` refuses it, in English.
 */
export type OutlineEditResult = { ok: true; rows?: Record<string, RowInput>; columns?: Record<string, ColumnInput> } | Refusal<'outline_deep' | 'not_grouped' | 'detail'>

/**
 * One structure operation (`editStructure`). Rows are 1-based numbers and columns are letters, as in A1
 * addresses; `count` is 1 by default; sheets and tables are matched case-insensitively. Tables move with their cells
 * and grow and shrink with the rows and columns inserted and deleted inside them; an inserted row gets the formulas of
 * the calculated columns, an inserted column a name (`Column1`), and references to a deleted column become #REF!.
 */
export type StructureOp =
  /** Inserts `count` rows before `row`: rows from `row` on move down; the new rows take the format of the row above. */
  | { op: 'insertRows'; sheet: string; row: number; count?: number }
  /** Deletes rows `row` … `row + count - 1`: the rows below move up. */
  | { op: 'deleteRows'; sheet: string; row: number; count?: number }
  /** Inserts `count` columns before `column`: columns from `column` on move right; the new ones take the format of the column on the left. */
  | { op: 'insertColumns'; sheet: string; column: string; count?: number }
  /** Deletes `count` columns from `column` on: the columns to the right move left. */
  | { op: 'deleteColumns'; sheet: string; column: string; count?: number }
  /** Cut and paste: the range moves so that its top-left cell lands on `to`, replacing what was there; cells keep their `meta`. */
  | { op: 'moveRange'; from: RangeRef; to: CellRef }
  /**
   * Copy and paste: the range is pasted at `to`; a target that is a whole multiple of the source's size is
   * tiled with copies (fill down and fill right are copies of the first row or column). Copies have no
   * `meta` unless `withMeta`.
   */
  | { op: 'copyRange'; from: RangeRef; to: RangeRef; withMeta?: boolean }
  /** A new name by Excel's rules (see `BookInput.sheets`) that no other sheet has. */
  | { op: 'renameSheet'; sheet: string; name: string }
  | { op: 'deleteSheet'; sheet: string }
  /**
   * A new empty sheet named `name` (by Excel's rules, not taken), before the sheet `before`, else after the last one.
   * A reference across sheets takes in a sheet inserted between its ends.
   */
  | { op: 'insertSheet'; name: string; before?: string }
  /**
   * The sheet moved before the sheet `before`, else after the last one; before a hidden sheet it goes before the next
   * visible one, as Excel moves it. A reference across sheets covers the sheets between its ends in the new order; an end
   * that moves stays its end until it passes the other end, then the next sheet inward takes its place.
   */
  | { op: 'moveSheet'; sheet: string; before?: string }
  /** The sheet hidden (Excel's Hide Sheet), very hidden with `very` (see `SheetInput.hidden`); a book keeps one visible sheet. */
  | { op: 'hideSheet'; sheet: string; very?: boolean }
  /** The sheet shown again (Excel's Unhide Sheet). */
  | { op: 'unhideSheet'; sheet: string }
  /** A table over cells of `sheet` (Excel's Format as Table): with a header row its header cells take the columns' names. */
  | { op: 'addTable'; sheet: string; table: TableInput }
  /** A new name for the table, by the rules of `TableInput.name`; every structured reference to it follows. */
  | { op: 'renameTable'; table: string; name: string }
  /** The table and its cells deleted (Excel's Delete Table): structured references to it become #REF!. */
  | { op: 'deleteTable'; table: string }
  /**
   * The table turned into plain cells (Excel's Convert to Range): structured references to it become references to its
   * cells with `$` and the sheet (`'S'!$B$2:$B$9`; a reference to the formula's own row keeps its row relative, `'S'!$B2`).
   */
  | { op: 'convertTable'; table: string }
  /**
   * Excel's Group and Ungroup of rows `5:7` or columns `B:D` (one line: `5`, `C`). Group puts each line one outline level
   * deeper, those at the deepest level (7) staying there, and is refused (`outline_deep`) when all of them are; lines next
   * to each other at a level make one group, so a group added right after another of its level joins it
   * (`outlineGroups`). Ungroup takes a level off each line that has one, refused (`not_grouped`) when none has. Hidden
   * lines stay hidden.
   */
  | { op: 'group' | 'ungroup'; sheet: string; rows: string }
  | { op: 'group' | 'ungroup'; sheet: string; columns: string }
  /**
   * Excel's Show Detail and Hide Detail, a group's `+` and `−` button: at a summary line (`OutlineGroup.summary`) on the
   * lines it sums up, at a line of a group on the group of the line's level that holds it, through its summary line. Hiding
   * hides the lines and marks the summary line `collapsed`; showing shows them but the lines of the collapsed groups inside,
   * and drops the mark. Refused at a line in no group or in a group with no summary line, past the sheet's edge
   * (`not_grouped`), and where the detail is shown already, or hidden (`detail`).
   */
  | { op: 'showDetail' | 'hideDetail'; sheet: string; row: number }
  | { op: 'showDetail' | 'hideDetail'; sheet: string; column: string }
  /**
   * Excel's level buttons (`Outline.ShowLevels`): the lines of outline level `rows` (or `columns`) and deeper hidden, the
   * other grouped lines shown, the collapsed marks as Excel keeps them; 8 shows every group. Lines outside groups stay as
   * they are.
   */
  | { op: 'showLevels'; sheet: string; rows?: number; columns?: number }
  /**
   * Excel's Merge Cells over `range` (`B2:D3`, whole columns and rows too): the range, grown by the merged areas it
   * meets, becomes one merged area (`SheetInput.merges`) that keeps its top-left cell; when that cell is empty, the
   * area's first cell holding a value or a formula, row by row, moves into it as cut and paste moves it. The other cells
   * lose their values and formulas and take the top-left cell's style and number format. With `across`, each row of the
   * range is merged on its own (Merge Across); with `center`, the area is centred as well (Merge & Center). A range that
   * meets a table is refused (`table`), and one that meets an array formula of the old kind (`array`).
   */
  | { op: 'mergeCells'; sheet: string; range: string; across?: boolean; center?: boolean }
  /** Excel's Unmerge Cells: the merged areas that meet `range` are split back into their cells. */
  | { op: 'unmergeCells'; sheet: string; range: string }
  /**
   * Excel's fill handle (`Range.AutoFill`): the cells of `range` go on over `to`, which holds `range` and reaches past it
   * along one axis, down, up, right or left (`A1:B2` over `A1:B9`, or over `A1:F2`). The cells added take the source's
   * cells in turn, as copies do (formulas, styles, number formats), the part of the source nearest to it at the far end
   * when `to` is not a whole number of sources long; refusals are a copy's. Each column of a fill down or up, and each row
   * of one right or left, goes on by itself, and what goes on as a series takes its next values instead (`FillType`,
   * docs/structure.md "The fill handle"). A `to` of another shape is an `InputError`.
   */
  | { op: 'fillRange'; sheet: string; range: string; to: string; type?: FillType }

/**
 * What a fill takes from its source (`StructureOp` `fillRange`), as Excel's `XlAutoFillType`. `default` is the drag of
 * the fill handle (`xlFillDefault`). A line's source is cut into runs of cells of one kind, empty cells skipped: numbers
 * go on by their step or trend; dates by days, by their step, or by months when they share a day of the month or are
 * months' ends; times by hours or their step; text with a number at the start (and a space) or at the end by the number,
 * quarters and en-US ordinals too; names of days and months through their list; dates, times and text without one step
 * are copied, and so are text, booleans and formulas. A single number goes on by 1 unless the source holds plain numbers
 * alone. `series` (`xlFillSeries`, the drag with Ctrl over a number) makes a single number go on by 1 always. `copy`
 * (`xlFillCopy`) copies every cell.
 */
export type FillType = 'default' | 'series' | 'copy'

/** Why Excel would refuse an operation. */
export type StructureRefusal =
  /** Cells would leave the sheet (past XFD or row 1,048,576), or a paste does not fit. */
  | 'off_sheet'
  /** A new sheet name that breaks Excel's rules or is taken (sheet names compare case-insensitively). */
  | 'sheet_name'
  /** A book keeps one visible sheet at least: its last visible sheet cannot be deleted or hidden. */
  | 'last_sheet'
  /** A very hidden sheet (`SheetInput.hidden` 'very') cannot be deleted or moved, as Excel refuses: it is shown again first. */
  | 'very_hidden'
  /** A move or copy would cut through a merged area, at its source or where it lands. */
  | 'merged'
  /** A copy that leaves out rows a filter hid, pasted over hidden rows of a filtered sheet (several areas onto several areas). */
  | 'filtered'
  /** The operation would change part of the area of an array formula of the old kind (`CellInput.array`). */
  | 'array'
  /**
   * The operation would shift or cut part of a table as Excel refuses it: rows or columns inserted into or deleted from
   * two tables at once, a table's header row deleted without the rest, a move that cuts a table, a table over another.
   */
  | 'table'
  /** A table name that breaks the rules of `TableInput.name` or that a table or a defined name of the book has. */
  | 'table_name'
  /** Group of lines that are all at the deepest outline level, 7. */
  | 'outline_deep'
  /** Ungroup of lines none of which has an outline level; Show or Hide Detail at a line in no group, or in a group with no summary line. */
  | 'not_grouped'
  /** Show Detail of a group shown already, Hide Detail of one hidden already (its summary line marked `collapsed`). */
  | 'detail'

/** `editStructure`'s answer. */
export type StructureResult =
  | {
      ok: true
      /** The new book; the input book is not changed, unchanged sheets and cells are shared. */
      book: Book
      /** Where a cell of the old book is now; null if it was deleted or overwritten, or its sheet is none of the old book's. */
      mapCell(ref: CellRef): CellRef | null
    }
  /** A refusal: its code, and a message in the language of the book's locale (English for en-US). */
  | Refusal<StructureRefusal>

/**
 * `Session.editStructure`: the edit's result and where the cells went, without the book (`getBook()` gives it), in
 * every kind of session. `changes` are the cells whose value or formula differs from those of the cell that was at
 * their old place (`mapCell`), by new address, a cell no old cell came to counting as empty before; with `display`, the
 * cells whose shown text or colour differs too. An interface that moves its shown cells by `moved` (or `mapCell`) and
 * applies these has the new ones.
 */
export type SessionStructureResult =
  | (EditResult & {
      ok: true
      /** Where a cell of the book before the operation is now; null if it was deleted or overwritten, or its sheet is none of that book's. */
      mapCell(ref: CellRef): CellRef | null
    })
  | Refusal<StructureRefusal>
