/** Structure operations (`editStructure`) and their answers, and a sheet's outline groups and commands over its rows and columns (`outlineGroups`, `outlineEdit`). */
import type { CellRef, LineSpan, RangeRef, Refusal } from './values.js'
import type { Book, ColumnInput, RowInput, TableInput } from './book.js'
import type { EditResult } from './edits.js'

/**
 * A group of rows or columns as Excel draws it (`outlineGroups`): the longest run of lines from `first` to `last`
 * (columns by number) at outline level `level` or deeper, so a group of level 2 lies inside one of level 1.
 */
export interface OutlineGroup extends LineSpan {
  level: number
  /**
   * The summary line, which holds the group's button: the line just after the group, below or to the right (just before
   * it with summaries above or to the left, `OutlineSettings`); null when that line is beyond the sheet's edge, where
   * Excel draws no button. Groups that end together share one.
   */
  summary: number | null
  /** Collapsed, with the button showing `+`: its summary line carries the collapsed marker as Excel keeps it (`RowInput.collapsed`); without a summary line, all its lines are hidden. */
  collapsed: boolean
}

/** A sheet's groups (`outlineGroups`), on each axis by first line, with the outer group before the inner one when two start together. */
export interface SheetOutline {
  rows: OutlineGroup[]
  columns: OutlineGroup[]
}

/** An outline operation of `StructureOp` without its sheet, applied to a single sheet (`outlineEdit`): `{ op: 'group', rows: '5:7' }`. */
export type OutlineEditOp = WithoutSheet<Extract<StructureOp, { op: 'group' | 'ungroup' | 'showDetail' | 'hideDetail' | 'showLevels' }>>

/** Each operation of a union without its `sheet`. */
type WithoutSheet<T> = T extends unknown ? Omit<T, 'sheet'> : never

/**
 * A sheet's rows and columns after an outline command (`outlineEdit`), each axis in full as `SheetInput` holds it, and
 * omitted when it has no entries; or why Excel rejects the command, as `editStructure` rejects it, in English.
 */
export type OutlineEditResult = { ok: true; rows?: Record<string, RowInput>; columns?: Record<string, ColumnInput> } | Refusal<'outline_deep' | 'not_grouped' | 'detail'>

/**
 * One structure operation (`editStructure`). Rows are 1-based numbers and columns are letters, as in A1
 * addresses; `count` is 1 by default; sheets and tables are matched case-insensitively. Tables move with their cells
 * and grow and shrink with the rows and columns inserted and deleted inside them; an inserted row gets the formulas of
 * the calculated columns, an inserted column a name (`Column1`), and references to a deleted column become #REF!.
 */
export type StructureOp =
  /** Inserts `count` rows before `row`: the rows from `row` onward move down; the new rows take the format of the row above. */
  | { op: 'insertRows'; sheet: string; row: number; count?: number }
  /** Deletes rows `row` … `row + count - 1`: the rows below move up. */
  | { op: 'deleteRows'; sheet: string; row: number; count?: number }
  /** Inserts `count` columns before `column`: the columns from `column` onward move right; the new ones take the format of the column to the left. */
  | { op: 'insertColumns'; sheet: string; column: string; count?: number }
  /** Deletes `count` columns starting at `column`: the columns to the right move left. */
  | { op: 'deleteColumns'; sheet: string; column: string; count?: number }
  /** Cut and paste: the range moves so that its top-left cell lands on `to`, replacing what was there; cells keep their `meta`. */
  | { op: 'moveRange'; from: RangeRef; to: CellRef }
  /**
   * Copy and paste: the range is pasted at `to`; a target that is a whole multiple of the source's size is
   * tiled with copies (fill down and fill right are copies of the first row or column). Copies have no
   * `meta` unless `withMeta`. `paste` pastes part of the copy, as Excel's Paste Special (`PasteKind`), `transpose`
   * turns it, rows into columns, `operation` combines the copy's numbers with the destination's, and `skipBlanks`
   * leaves the contents and look of the destination's cells under blank cells of the copy, and their notes under cells
   * without a note (`docs/structure.md` "Paste Special"); `operation` or `skipBlanks` without `paste` is Paste
   * Special's `all`.
   */
  | { op: 'copyRange'; from: RangeRef; to: RangeRef; withMeta?: boolean; paste?: PasteKind; transpose?: boolean; operation?: PasteOperation; skipBlanks?: boolean }
  /** A new name that follows Excel's rules (see `BookInput.sheets`) and is not used by another sheet. */
  | { op: 'renameSheet'; sheet: string; name: string }
  | { op: 'deleteSheet'; sheet: string }
  /**
   * Inserts a new empty sheet named `name` (following Excel's rules, not already used), before the sheet `before`, or
   * else after the last one. A reference that spans sheets includes a sheet inserted between its ends.
   */
  | { op: 'insertSheet'; name: string; before?: string }
  /**
   * Moves the sheet before the sheet `before`, or else after the last one; when `before` is hidden, the sheet goes before
   * the next visible one, as Excel moves it. A reference that spans sheets covers the sheets between its ends in the new
   * order; an end sheet that moves remains an end until it passes the other end, and then the next sheet inward takes
   * its place.
   */
  | { op: 'moveSheet'; sheet: string; before?: string }
  /** Hides the sheet (Excel's Hide Sheet), or makes it very hidden with `very` (see `SheetInput.hidden`); a workbook always keeps one visible sheet. */
  | { op: 'hideSheet'; sheet: string; very?: boolean }
  /** Unhides the sheet (Excel's Unhide Sheet). */
  | { op: 'unhideSheet'; sheet: string }
  /** Creates a table over cells of `sheet` (Excel's Format as Table): with a header row, its header cells take the column names. */
  | { op: 'addTable'; sheet: string; table: TableInput }
  /** A new name for the table, following the rules of `TableInput.name`; every structured reference to it is updated. */
  | { op: 'renameTable'; table: string; name: string }
  /** Deletes the table and its cells (Excel's Delete Table): structured references to it become #REF!. */
  | { op: 'deleteTable'; table: string }
  /**
   * Converts the table to plain cells (Excel's Convert to Range): structured references to it become references to its
   * cells with `$` and the sheet name (`'S'!$B$2:$B$9`; a reference to the formula's own row keeps the row relative,
   * `'S'!$B2`).
   */
  | { op: 'convertTable'; table: string }
  /**
   * Excel's Group and Ungroup of rows `5:7` or columns `B:D` (a single line: `5`, `C`). Group moves each line one
   * outline level deeper, leaving those already at the deepest level (7) there, and is rejected (`outline_deep`) when
   * all of them are at that level; adjacent lines at a level form one group, so a group added right after another of
   * its level merges with it (`outlineGroups`). Ungroup removes one level from each line that has one, and is rejected
   * (`not_grouped`) when none has. Hidden lines stay hidden.
   */
  | { op: 'group' | 'ungroup'; sheet: string; rows: string }
  | { op: 'group' | 'ungroup'; sheet: string; columns: string }
  /**
   * Excel's Show Detail and Hide Detail, a group's `+` and `−` buttons: at a summary line (`OutlineGroup.summary`), on
   * the lines it summarizes; at a line inside a group, on the group at that line's level that contains it, by way of the
   * group's summary line. Hiding hides the lines and marks the summary line `collapsed`; showing unhides them, except the
   * lines of the collapsed groups inside, and removes the marker. Rejected at a line in no group, or in a group whose
   * summary line is beyond the sheet's edge (`not_grouped`), and where the detail is already shown, or already hidden
   * (`detail`).
   */
  | { op: 'showDetail' | 'hideDetail'; sheet: string; row: number }
  | { op: 'showDetail' | 'hideDetail'; sheet: string; column: string }
  /**
   * Excel's outline level buttons (`Outline.ShowLevels`): hides the lines at outline level `rows` (or `columns`) and
   * deeper, shows the other grouped lines, and sets the collapsed markers as Excel keeps them; 8 shows every group. Lines
   * outside groups stay as they are.
   */
  | { op: 'showLevels'; sheet: string; rows?: number; columns?: number }
  /**
   * Excel's Merge Cells over `range` (`B2:D3`, including entire columns and rows): the range, extended by the merged
   * areas it overlaps, becomes one merged area (`SheetInput.merges`) that keeps its top-left cell; when that cell is
   * empty, the area's first cell holding a value or a formula, row by row, moves into it as cut and paste would move it.
   * The other cells lose their values and formulas and take the top-left cell's style and number format. With `across`,
   * each row of the range is merged separately (Merge Across); with `center`, the area is also centered (Merge &
   * Center). A range that overlaps a table is rejected (`table`), as is one that overlaps a legacy array formula
   * (`array`).
   */
  | { op: 'mergeCells'; sheet: string; range: string; across?: boolean; center?: boolean }
  /** Excel's Unmerge Cells: the merged areas that overlap `range` are split back into their cells. */
  | { op: 'unmergeCells'; sheet: string; range: string }
  /**
   * Excel's fill handle (`Range.AutoFill`): the cells of `range` are extended over `to`, which contains `range` and
   * extends beyond it along one axis, down, up, right or left (`A1:B2` over `A1:B9`, or over `A1:F2`). The added cells
   * take the source's cells in turn, as copies do (formulas, styles, number formats), with the part of the source
   * nearest to it at the far end when `to` is not a whole number of source lengths; the rejections are those of a copy.
   * Each column of a fill down or up, and each row of a fill right or left, is extended independently, and whatever
   * continues as a series takes its next values instead (`FillType`, docs/structure.md "The fill handle"). A `to` of any
   * other shape throws an `InputError`.
   */
  | { op: 'fillRange'; sheet: string; range: string; to: string; type?: FillType }

/**
 * What a fill takes from its source (`StructureOp` `fillRange`), like Excel's `XlAutoFillType`. `default` is dragging
 * the fill handle (`xlFillDefault`). Each line's source is split into runs of cells of the same kind, skipping empty
 * cells: numbers continue by their step or trend; dates by days, by their step, or by months when they share a day of
 * the month or are month ends; times by hours or by their step; text with a number at the start (followed by a space)
 * or at the end continues by the number, including quarters and en-US ordinals; names of days and months continue
 * through their lists; dates, times and text without a constant step are copied, as are plain text, booleans and
 * formulas. A single number continues by 1 unless the source holds only plain numbers. `series` (`xlFillSeries`,
 * dragging with Ctrl over a number) always continues a single number by 1. `copy` (`xlFillCopy`) copies every cell.
 */
export type FillType = 'default' | 'series' | 'copy'

/**
 * What a copy pastes of its source (`StructureOp` `copyRange`), as Excel's Paste Special. `all`, the default, pastes
 * everything. `values`: each cell's value as computed, a formula's and a spilled cell's too, an error as an error
 * constant; `formulas`: the formulas as a copy pastes them and the constants. Both leave the target's styles, number
 * formats, merged areas, notes, hyperlinks, conditional formats and data validation, a blank of the source clears the
 * target's cell, and both are refused onto a merged area the copy does not bring. `formats`: styles, number formats,
 * merged areas and conditional formats, the target's contents, notes, hyperlinks and data validation left.
 * `validation`: the data validation alone. `valuesAndNumberFormats` and `formulasAndNumberFormats`: as `values` and
 * `formulas` with the copy's number formats, and not refused onto merged areas. `columnWidths`: the widths of the
 * copy's columns alone. `allExceptBorders`: as `all`, each cell keeping the destination's borders. `notes`: the copy's
 * notes alone, a cell without one clearing the destination's (Excel's Comments). `allMergingConditionalFormats`: as
 * `all`, the destination's conditional formats kept whole after the copy's, a rule copied on its own sheet growing
 * where it makes one rectangle with the pasted cells; with `skipBlanks` it is `all`.
 * Excel's Format Painter is `formats`; its All Using Source Theme is `all` within one workbook.
 */
export type PasteKind = 'all' | 'values' | 'formulas' | 'formats' | 'validation' | 'valuesAndNumberFormats' | 'formulasAndNumberFormats' | 'columnWidths' | 'allExceptBorders' | 'notes' | 'allMergingConditionalFormats'

/**
 * Paste Special's operation (`StructureOp` `copyRange`): each number of the copy is added to, subtracted from,
 * multiplied into or divided into the destination cell's, a blank counting as 0; a formula on either side makes a
 * formula of both (`10+(A1*2)`). Text, booleans and errors of the copy leave the destination's cell. It applies to the
 * kinds that paste contents and is ignored by the others; onto an array spilled over more than one cell it is refused
 * (`array`).
 */
export type PasteOperation = 'add' | 'subtract' | 'multiply' | 'divide'

/** Why Excel would reject an operation. */
export type StructureRefusal =
  /** Cells would go off the sheet (beyond XFD or row 1,048,576), or a paste does not fit. */
  | 'off_sheet'
  /** A new sheet name that breaks Excel's rules or is already used (sheet names are compared case-insensitively). */
  | 'sheet_name'
  /** A workbook keeps at least one visible sheet: its last visible sheet cannot be deleted or hidden. */
  | 'last_sheet'
  /** A very hidden sheet (`SheetInput.hidden` 'very') cannot be deleted or moved, as in Excel: unhide it first. */
  | 'very_hidden'
  /** A move or copy would split a merged area, at its source or at its destination. */
  | 'merged'
  /** A copy that skips rows hidden by a filter, pasted over hidden rows of a filtered sheet (multiple areas onto multiple areas). */
  | 'filtered'
  /** The operation would change part of a legacy array formula's range (`CellInput.array`), or Paste Special's operation a spilled array. */
  | 'array'
  /**
   * The operation would shift or split part of a table in a way Excel rejects: rows or columns inserted into or deleted
   * from two tables at once, a table's header row deleted without the rest of the table, a move that splits a table, a
   * table on top of another.
   */
  | 'table'
  /** A table name that breaks the rules of `TableInput.name` or is already used by a table or a defined name in the workbook. */
  | 'table_name'
  /** Group on lines that are all at the deepest outline level, 7. */
  | 'outline_deep'
  /** Ungroup on lines none of which has an outline level; Show or Hide Detail at a line in no group, or in a group without a summary line. */
  | 'not_grouped'
  /** Show Detail on a group already shown, Hide Detail on one already hidden (its summary line marked `collapsed`). */
  | 'detail'

/** The result of `editStructure`. */
export type StructureResult =
  | {
      ok: true
      /** The new workbook; the input workbook is not modified, and unchanged sheets and cells are shared. */
      book: Book
      /** Where a cell of the old workbook is now; null if it was deleted or overwritten, or if its sheet is not in the old workbook. */
      mapCell(ref: CellRef): CellRef | null
    }
  /** A rejection: its code, and a message in the language of the workbook's locale (English for en-US). */
  | Refusal<StructureRefusal>

/**
 * `Session.editStructure`: the edit's result and where the cells went, without the workbook (`getBook()` returns it),
 * in every kind of session. `changes` are the cells whose value or formula differs from that of the cell previously at
 * their old position (`mapCell`), by new address, with a cell that no old cell moved to counting as empty before; with
 * `display`, also the cells whose displayed text or color differs. An interface that moves its displayed cells
 * according to `moved` (or `mapCell`) and applies these changes has the new cells.
 */
export type SessionStructureResult =
  | (EditResult & {
      ok: true
      /** Where a cell of the workbook before the operation is now; null if it was deleted or overwritten, or if its sheet is not in that workbook. */
      mapCell(ref: CellRef): CellRef | null
    })
  | Refusal<StructureRefusal>
