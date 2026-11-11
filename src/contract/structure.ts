/** Structure operations (`editStructure`) and their answers, and a sheet's outline groups and commands over its rows and columns (`outlineGroups`, `outlineEdit`). */
import type { CellRef, LineSpan, RangeRef, Refusal } from './values.js'
import type { Book, ColumnInput, FilterCriteria, FilterSort, RowInput, TableInput } from './book.js'
import type { PivotGrouping, PivotRefusalReason, PivotRef, PivotSource, PivotTablePatch, Slicer, SlicerCachePatch, SlicerPatch, Timeline, TimelineCachePatch, TimelinePatch } from './pivot.js'
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
   * Excel's Insert PivotTable: a PivotTable named `name` (the first free `PivotTable1`, `PivotTable2`… without it) at
   * `cell` of `sheet`, or on a new sheet named `sheet` inserted as `insertSheet` inserts it (`newSheet`, at A3 without
   * `cell`). A source by range, table or name makes a new pivot cache of its fields, its header row's; `{ cache }` shares
   * a cache of the workbook by its `id`. `definition` lays out its fields as `editPivotTable` does, `groupDates` too.
   * With `xlcalc/pivot`, a PivotTable laid over cells that hold something or over a table is made as Excel 365 makes it:
   * it shows `#SPILL!` in its top-left cell, the cells in the way left as they are, until a refresh finds the place
   * free; a value in its top-left cell it takes over. Refused (`pivot`) over another PivotTable, with a name already
   * used on the sheet, and, without `xlcalc/pivot`, over cells that are not empty or a table.
   */
  | { op: 'addPivotTable'; sheet: string; cell?: string; newSheet?: { before?: string }; name?: string; source: PivotSource | { cache: number }; definition?: PivotTablePatch; groupDates?: boolean }
  /**
   * Changes a PivotTable's definition (`PivotTablePatch`): fields on its axes and their order, value fields, items hidden
   * or collapsed, filters, sorting, totals, layout, captions, style, options; `source` is Excel's Change Data Source, a
   * new cache for this PivotTable alone when its cache is shared. With `xlcalc/pivot` loaded the PivotTable is
   * refreshed; without it, its cells stay and its cache is marked `stale` (`StructureResult.stale`). `groupDates`: a
   * field of dates the edit puts on the rows or the columns is grouped as Excel 365 groups it when it is checked in the
   * PivotTable Fields list (time grouping), as `groupPivotField` groups it; Excel does not group it when a macro places
   * it, nor without the flag here. It needs `xlcalc/pivot`.
   */
  | { op: 'editPivotTable'; sheet: string; pivot: string; set: PivotTablePatch; source?: PivotSource; groupDates?: boolean }
  /** A new name for the PivotTable, unique among its sheet's; slicers and timelines that name it follow. */
  | { op: 'renamePivotTable'; sheet: string; pivot: string; name: string }
  /** Excel's Move PivotTable: the PivotTable's cells move so that its top-left cell (its filters' when it has them) lands on `to`, on its sheet or another. */
  | { op: 'movePivotTable'; sheet: string; pivot: string; to: CellRef }
  /** Deletes the PivotTable and clears its cells; its cache goes with the last PivotTable and slicer over it. */
  | { op: 'deletePivotTable'; sheet: string; pivot: string }
  /**
   * Excel's Refresh: each cache of the PivotTables `pivots`, and every PivotTable over it, from its source; without
   * `pivots`, every cache of the workbook that can be refreshed (Refresh All), the others skipped. Needs `xlcalc/pivot`;
   * a PivotTable that is not there, a source that cannot be refreshed, or the subpath not loaded, is refused (`pivot`).
   */
  | { op: 'refreshPivots'; pivots?: PivotRef[] }
  /**
   * Excel's Group on the field `field` (its index in the cache) of a PivotTable, for every PivotTable of its cache, laid
   * out again in one step. Refused (`pivot`) where Excel refuses: text among the dates or numbers grouped, a field with
   * calculated items, items by hand of a field of dates or numbers grouped; and without `xlcalc/pivot`.
   */
  | { op: 'groupPivotField'; sheet: string; pivot: string; field: number; group: PivotGrouping }
  /**
   * Excel's Ungroup: a field of dates or numbers back to its values, the fields its periods made gone from every
   * PivotTable of the cache; a field of groups made by hand gone; with `items` (indexes into the PivotTable field's
   * `items`), those groups by hand alone, their items each a group of its own again.
   */
  | { op: 'ungroupPivotField'; sheet: string; pivot: string; field: number; items?: number[] }
  /** A group's label in the cache (`Group1` to `East and West`): the item `item` of the field `field`, on every PivotTable of the cache. */
  | { op: 'renamePivotGroup'; sheet: string; pivot: string; field: number; item: number; name: string }
  /**
   * Excel's Insert Calculated Field: the cache's calculated field `name`, a formula over fields by name as Excel's dialog
   * takes it (`=Amount/Qty`, `='Unit Price'*Qty`), added, or its formula changed when the cache has it (Modify); a null
   * `formula` deletes it, and it leaves every PivotTable of the cache. Added with `values`, it goes on the Values of
   * the PivotTable `pivot`, last, with Sum and the name Excel's dialog gives it (`Sum of Unit`), in the same step;
   * without, it is placed by `editPivotTable`, as Excel's object model adds it.
   */
  | { op: 'setPivotCalculatedField'; sheet: string; pivot: string; name: string; formula: string | null; values?: boolean }
  /**
   * Excel's Insert Calculated Item of the field `field`: the item `name`, a formula over the field's items as Excel's
   * dialog takes it (`=East+West`, `='North'*2`), added on every PivotTable of the cache, or its formula changed when
   * the field has it; a null `formula` deletes it.
   */
  | { op: 'setPivotCalculatedItem'; sheet: string; pivot: string; field: number; name: string; formula: string | null }
  /**
   * Excel's Insert Slicer on `sheet`: over a slicer cache of the workbook by its name, or a new one of the source field
   * `field` that filters the PivotTables `pivots`, all of one pivot cache, or a new one of a table's column (`table` and
   * `column` by their names), which chooses by the table's filter (`applyFilter`). Its name is the first free from the
   * field's (`Region`, `Region 1`…) without `slicer.name`, and a new cache's `Slicer_Region`. A table's slicer and its
   * cache go with the table or the column (`deleteTable`, `convertTable`, `deleteColumns`, `deleteSheet`), as in Excel.
   */
  | { op: 'addSlicer'; sheet: string; slicer: Omit<Slicer, 'name' | 'cache'> & { name?: string }; cache: string | { field: string; pivots: PivotRef[] } | { table: string; column: string } }
  /** Changes a slicer (`set`) and its cache (`cache`): the items chosen, the PivotTables it filters (Report Connections), its name, look and place. */
  | { op: 'editSlicer'; slicer: string; set?: SlicerPatch; cache?: SlicerCachePatch }
  /** Deletes the slicer; its cache goes with its last slicer. */
  | { op: 'deleteSlicer'; slicer: string }
  /**
   * Excel's Insert Timeline, as `addSlicer`, of a date field. Another timeline over the same field of a shared PivotTable reuses its cache and choice; new connections join it.
   */
  | { op: 'addTimeline'; sheet: string; timeline: Omit<Timeline, 'name' | 'cache'> & { name?: string }; cache: string | { field: string; pivots: PivotRef[] } }
  /** Changes a timeline and its cache: the dates chosen, the PivotTables it filters, its name, level, look and place. */
  | { op: 'editTimeline'; timeline: string; set?: TimelinePatch; cache?: TimelineCachePatch }
  /** Deletes the timeline; its cache goes with its last timeline. */
  | { op: 'deleteTimeline'; timeline: string }
  /**
   * Excel's filter, of the sheet's AutoFilter (`sheet`) or of a table (`table`). With `criteria`, they replace the
   * filter's (a column's button, `Range.AutoFilter` with a field); without, those it holds are applied again (Data >
   * Reapply). Each row of its range below the header that every column's condition keeps is shown, one hidden by hand
   * too, and the others are hidden; then its sort, if it has one, sorts the rows shown as `sortRange` does, the hidden
   * ones staying in place. The sheet's AutoFilter filters the range it acts on (`Session.getAutoFilterRange`) and takes
   * it as its own; a table's totals row and a last row of `SUBTOTAL` under an AutoFilter stay out. Top, average and
   * relative date filters keep the bounds they found (`TopFilter.filterValue`, `DynamicFilter.value` and `maxValue`),
   * relative dates counted from the day of the calculation, and the criteria take the form Excel saves after it: the
   * chosen values in order of their text and the dates in time order, a number or date operand of a custom condition as
   * its number, equality alone without wildcards as a list of values. The sheet is `filtered` while a filter of it has a
   * condition. Refused (`filter`) where a top or average filter's column holds an error, as Excel refuses it. A sheet
   * without an AutoFilter, or a table without filter buttons (`TableInput.filter` false), is an input error.
   */
  | { op: 'applyFilter'; sheet: string; criteria?: FilterCriteria }
  | { op: 'applyFilter'; table: string; criteria?: FilterCriteria }
  /**
   * Excel's Clear (`ShowAllData`): the conditions of the filter's columns and its sort removed, their buttons kept, and
   * the rows of its range below the header shown, rows hidden by hand too; the cells stay where they are. With `column`
   * (from 0 at the filter's left), that column's condition alone (Clear Filter From), the others and the sort applied
   * again as `applyFilter` applies them.
   */
  | { op: 'clearFilter'; sheet: string; column?: number }
  | { op: 'clearFilter'; table: string; column?: number }
  /**
   * Excel's Sort of a range (Data > Sort, `Range.Sort`): the rows of `sort.range` reordered by its conditions in turn
   * (its columns, with `byColumns`), each line's cells moving together, formulas as a move moves them, and lines hidden
   * staying in place. Ties keep their order. The sort is kept as the sheet's `sort`, as Excel's `Worksheet.Sort` keeps
   * it over a filter's data too; a filter's own sort is set by `applyFilter`'s criteria. Refused over merged areas of
   * different sizes (`merged`), part of a legacy array formula's range (`array`), part of a table (`table`) or a
   * PivotTable (`pivot`).
   */
  | { op: 'sortRange'; sheet: string; sort: FilterSort }

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
 * kinds that paste contents and is ignored by the others; onto an array spilled over more than one cell or a legacy
 * array formula, of one cell or its whole range, it is refused (`array`).
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
  /**
   * The operation would cut a PivotTable or lay one over cells that are not empty, a table or another PivotTable; a
   * PivotTable's, slicer's or timeline's name breaks the rules or is used; a field not of the cache; a refresh that cannot
   * be made (its source, or `xlcalc/pivot` not loaded).
   */
  | 'pivot'
  /** A top or average filter over a column that holds an error (`applyFilter`, `clearFilter`). */
  | 'filter'

/** The result of `editStructure`. */
export type StructureResult =
  | {
      ok: true
      /** The new workbook; the input workbook is not modified, and unchanged sheets and cells are shared. */
      book: Book
      /** Where a cell of the old workbook is now; null if it was deleted or overwritten, or if its sheet is not in the old workbook. */
      mapCell(ref: CellRef): CellRef | null
      /** The PivotTables whose definition changed and whose cells were not laid out again, `xlcalc/pivot` not loaded (`PivotCache.stale`). */
      stale?: PivotRef[]
    }
  /** A rejection: its code, and a message in the language of the workbook's locale (English for en-US). */
  | StructureRefused

/** An operation `editStructure` rejected: its code and message; over a PivotTable's cells, Excel's alert (`reason`). */
export type StructureRefused = Refusal<StructureRefusal> & { reason?: PivotRefusalReason }

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
      /** As `StructureResult.stale`. */
      stale?: PivotRef[]
    })
  | StructureRefused
