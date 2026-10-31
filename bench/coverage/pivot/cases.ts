/**
 * Pivot tables, for Excel for Windows (`bench/windows/run.ts pivot <set>`): a set's cases describe pivot tables as data,
 * and `bench/windows/pivot.ts` turns each description into steps of the object model (`PivotCaches.Create`,
 * `CreatePivotTable`, the fields' `Orientation`, `AddDataField`, `PivotFilters.Add2`, `Range.Group`, `SlicerCaches.Add2`
 * and the rest). Excel builds each case's pivot tables in a book of its own over the rows of `data.ts`, refreshes and
 * saves it; the saved book and what Excel showed in each pivot table's area are the answer. A case's operations run in
 * phases: `read` ends one (each pivot table read, the cells watched, a copy of the book saved), so a second step over
 * the same pivot table (a field moved, a filter, an edit of the source and a refresh) has its own answer.
 *
 * Names follow the file's vocabulary where it has one (ECMA-376 `ST_DataConsolidateFunction`, `ST_ShowDataAs`).
 */
import type { CellValue } from '../../../src/contract.js'
import type { SourceKind } from './data.js'

/** A data field's function (`ST_DataConsolidateFunction`). */
export type Agg = 'sum' | 'count' | 'average' | 'max' | 'min' | 'product' | 'countNums' | 'stdDev' | 'stdDevp' | 'var' | 'varp'

export const AGGS: readonly Agg[] = ['sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp']

/** A data field's calculation, Show Values As (`ST_ShowDataAs`, with the ranks and the shares of a parent of Excel 2010). */
export type ShowAs =
  | 'normal'
  | 'difference'
  | 'percentDiff'
  | 'percent'
  | 'percentOfRow'
  | 'percentOfCol'
  | 'percentOfTotal'
  | 'index'
  | 'runTotal'
  | 'percentOfRunningTotal'
  | 'percentOfParentRow'
  | 'percentOfParentCol'
  | 'percentOfParent'
  | 'rankAscending'
  | 'rankDescending'

/** Where a field goes. */
export type Axis = 'row' | 'column' | 'page' | 'hidden'

/** A data field: the source field, its function (`sum` by default), its caption and number format, and its calculation over a base field and item (`(previous)`, `(next)` or an item's name). */
export interface ValueField {
  field: string
  fn?: Agg
  caption?: string
  format?: string
  show?: ShowAs
  base?: string
  item?: string
}

/** The pivot table's own settings. */
export interface TableProps {
  rowGrand?: boolean
  columnGrand?: boolean
  grandTotalName?: string
  /** The header of the rows in the compact form (`Row Labels`), of the columns (`Column Labels`). */
  rowHeader?: string
  columnHeader?: string
  /** What a cell with an error shows, what an empty cell shows. */
  errorText?: string
  emptyText?: string
  /** Several filters on one field (a label filter and a value filter). */
  multipleFilters?: boolean
  /** The row naming the data fields (`Values`) shown, as by default. */
  valuesRow?: boolean
  /** Items sorted by Excel's custom lists (the month and day names), as by default. */
  customLists?: boolean
  /** Merge and center cells with labels. */
  mergeLabels?: boolean
  /** Classic PivotTable layout (fields dragged in the grid, `InGridDropZones`). */
  classic?: boolean
  /** Report filter fields per column (or per row, across first): `PageFieldWrapCount`; across then down: `PageFieldOrder`. */
  pagesWrap?: number
  pagesAcross?: boolean
}

/** A field's settings. */
export interface FieldProps {
  caption?: string
  form?: 'compact' | 'outline' | 'tabular'
  blankLine?: boolean
  subtotalsTop?: boolean
  repeatLabels?: boolean
  format?: string
  showAllItems?: boolean
}

/** The filters of a field, `XlPivotFilterType` by its name without `xl`, and its number. */
export const FILTER_TYPES = {
  TopCount: 1, BottomCount: 2, TopPercent: 3, BottomPercent: 4, TopSum: 5, BottomSum: 6,
  ValueEquals: 7, ValueDoesNotEqual: 8, ValueIsGreaterThan: 9, ValueIsGreaterThanOrEqualTo: 10, ValueIsLessThan: 11, ValueIsLessThanOrEqualTo: 12, ValueIsBetween: 13, ValueIsNotBetween: 14,
  CaptionEquals: 15, CaptionDoesNotEqual: 16, CaptionBeginsWith: 17, CaptionDoesNotBeginWith: 18, CaptionEndsWith: 19, CaptionDoesNotEndWith: 20, CaptionContains: 21, CaptionDoesNotContain: 22,
  CaptionIsGreaterThan: 23, CaptionIsGreaterThanOrEqualTo: 24, CaptionIsLessThan: 25, CaptionIsLessThanOrEqualTo: 26, CaptionIsBetween: 27, CaptionIsNotBetween: 28,
  SpecificDate: 29, NotSpecificDate: 30, Before: 31, BeforeOrEqualTo: 32, After: 33, AfterOrEqualTo: 34, DateBetween: 35, DateNotBetween: 36,
  AllDatesInPeriodQuarter1: 53, AllDatesInPeriodQuarter2: 54, AllDatesInPeriodQuarter3: 55, AllDatesInPeriodQuarter4: 56,
  AllDatesInPeriodJanuary: 57, AllDatesInPeriodFebruary: 58, AllDatesInPeriodMarch: 59, AllDatesInPeriodApril: 60, AllDatesInPeriodMay: 61, AllDatesInPeriodJune: 62,
  AllDatesInPeriodJuly: 63, AllDatesInPeriodAugust: 64, AllDatesInPeriodSeptember: 65, AllDatesInPeriodOctober: 66, AllDatesInPeriodNovember: 67, AllDatesInPeriodDecember: 68,
} as const

export type FilterType = keyof typeof FILTER_TYPES

/** The periods of a date grouping, as Excel's Grouping dialog lists them. */
export type Period = 'seconds' | 'minutes' | 'hours' | 'days' | 'months' | 'quarters' | 'years'

/** A filter's or a group's value: a number, text, or a date as `{ d: serial }`. */
export type FilterValue = number | string | { d: number }

/** An edit of a cell (`D!A50`): a number or text set as its value (`Value2`), a formula, text typed as Excel 365 takes it (`Formula2`), or null to clear it. */
export type EditValue = null | number | string | { formula: string } | { typed: string }

/**
 * An operation of a case, on the pivot table `pt` (`PT1` by default). `PT1` is made first over the case's source at
 * `P!A3`; `create` makes another, over the cache of `PT1` unless it names a source.
 */
export type PivotOp = (
  | { create: string; at: string; source?: SourceKind }
  | { rows: string[] }
  | { columns: string[] }
  | { pages: string[] }
  | { move: string; to: Axis; position?: number }
  | { value: ValueField }
  /** A data field added without a function (`AddDataField` with the field alone): the function Excel gives it. */
  | { valueDefault: string }
  /** A data field by its caption taken off Values (`DataFields(caption).Orientation` hidden). */
  | { dropValue: string }
  /** Fields placed by `PivotTable.AddFields` (added to those there), or a field's `CubeField` given an orientation. */
  | { addFields: { rows?: string[]; columns?: string[] } }
  | { cubeField: string; to: 'row' | 'column' }
  /** A cell's number format (`P!H5`), cells merged (`P!H6:I6`), a table with headers over a range named `name`. */
  | { cellFormat: string; numberFormat: string }
  | { merge: string }
  | { addTable: string; name: string }
  /** The data fields' own field (Σ Values) on rows or columns, at a position. */
  | { valuesOn: 'row' | 'column'; position?: number }
  | { layout: 'compact' | 'outline' | 'tabular' }
  | { repeatAll: boolean }
  | { table: TableProps }
  | { field: string; props: FieldProps }
  /** A field's subtotals: Excel's own, none, or the functions listed. */
  | { subtotals: string; fns: Agg[] | 'automatic' | 'none' }
  | { hide: string; items: string[] }
  /** A page field's selection: one item (`CurrentPage`), several (multiple items, the rest hidden), or `(All)`. */
  | { page: string; items: string[] | '(All)' }
  /** A filter on a field; `by` names the data field (its caption) a value filter or top 10 looks at. */
  | { filter: string; type: FilterType; value1?: FilterValue; value2?: FilterValue; by?: string; wholeDays?: boolean }
  /** AutoSort of a field by its labels, or by a data field (its caption); `manual` turns sorting on each update off and keeps the order shown. */
  | { sort: string; order: 'ascending' | 'descending' | 'manual'; by?: string }
  /** An item moved by hand to a position. */
  | { item: string; of: string; position: number }
  | { calcField: string; formula: string }
  | { calcItem: string; field: string; formula: string }
  /** Excel's grouping of a field of dates or numbers: start and end (Excel's own where left out), an interval, the periods. */
  | { group: string; start?: FilterValue; end?: FilterValue; by?: number; periods?: Period[] }
  | { groupItems: string; items: string[] }
  /** Items collapsed (`ShowDetail` false), or every item of the field; `expand` shows them again. */
  | { collapse: string; items: string[] | 'all'; expand?: boolean }
  | { ungroup: string }
  /** A slicer on a field, on the sheet `on` (`P` by default), over the pivot tables named besides `pt`; where Excel puts it without a place (`auto`); the slicer named `slicerName` apart from its cache `name`. */
  | { slicer: string; name: string; pivots?: string[]; on?: string; auto?: boolean; slicerName?: string }
  /** A slicer's or a timeline's properties read by its name (`Slicer.Name`, `Top`), and each of its cache's items' (`SlicerItem.Caption`), as the operation's result. */
  | { slicerGet: string; get?: string[]; items?: string[] }
  /** A slicer's or a timeline's property set by its name (`Slicer.Name`, `Caption`, `Locked`, `DisableMoveResizeUI`, `Top`, `Width`). */
  | { slicerSet: string; prop: string; to: string | number | boolean }
  /** A timeline's dates set (`TimelineState.SetFilterDateRange`), by its cache's name. */
  | { timelineSet: string; from: number; to: number }
  /** The sheet protected without a password (`Worksheet.Protect`), Edit objects allowed (`objects`) or not, Use PivotTable reports (`pivots`) or not; or unprotected. */
  | { protect: string; objects?: boolean; pivots?: boolean }
  | { unprotect: string }
  /** A slicer's or timeline's cache deleted with its slicers (`SlicerCache.Delete`). */
  | { deleteSlicer: string }
  /** A sheet added at the end, or deleted with what is on it. */
  | { addSheet: string }
  | { deleteSheet: string }
  /** A slicer's items selected, the others not. */
  | { select: string; items: string[] }
  /** A slicer's order of items and what it does with items that have no data (`XlSlicerCrossFilterType`: none, at the end, dimmed in place, hidden). */
  | { slicerSettings: string; order?: 'ascending' | 'descending' | 'source'; noData?: 'none' | 'last' | 'dimmed' | 'hidden' }
  /** A pivot table connected to a slicer, or disconnected (`off`). */
  | { link: string; pivot: string; off?: boolean }
  /** A timeline on a date field, on the sheet `on` (`P` by default), its range of dates and its level; where Excel puts it without a place (`auto`). */
  | { timeline: string; name: string; from?: number; to?: number; level?: 'years' | 'quarters' | 'months' | 'days'; on?: string; auto?: boolean }
  | { cache: { missingItems?: 'default' | 'none' | 'max'; refreshOnLoad?: boolean } }
  /** A cell of the source edited, or a formula typed on `G` (as Excel 365 types it, `Formula2`). */
  | { edit: string; to: EditValue }
  /** Show Details for a value cell; its new sheet's source records are an operation result. */
  | { drill: string }
  | { refresh: true }
  /** A defined name given another reference (`D!$A$1:$J$30`), the table `Sales` resized to a range of `D`. */
  | { define: string; ref: string }
  | { resize: string; range: string }
  /** Every pivot table of the book refreshed (`RefreshTable` each). */
  | { refreshAll: true }
  /**
   * An operation of the structure over `P` (`Struct` of `xl.ps1`): rows or columns inserted or deleted at `range`, the
   * range copied or cut (`moveRange`) to `to` on the sheet `target`, filled down over `to` (`FillDown`), merged, made a
   * table named `target`.
   */
  | { struct: 'insertRows' | 'deleteRows' | 'insertColumns' | 'deleteColumns' | 'copyRange' | 'moveRange' | 'fillDown' | 'mergeCells' | 'addTable'; range?: string; to?: string; target?: string; sheet?: string }
  /** Text put on the clipboard and pasted at a cell of `P` (`Set-Clipboard`, `Worksheet.Paste`). */
  | { paste: string; at: string }
  /** A range copied (`Range.Copy`, `D!A2`) and pasted at a cell of `P` with it selected (`Worksheet.Paste`), as Ctrl+C and Ctrl+V. */
  | { copyPaste: string; at: string }
  /** A range cleared (`P!A1:D10`, areas apart by commas): everything (`Range.Clear`) or its formats (`ClearFormats`). */
  | { clear: string; what: 'all' | 'formats' }
  /** The end of a phase: each pivot table read, the cells watched, a copy of the book saved. */
  | { read: true }
  /**
   * Formulas over a pivot table in another book (`book`, a saved book of a set, `basic/sum`), opened beside as
   * `Src.xlsx`: typed with it open and read; read again with it closed; the book saved, opened again alone without
   * updating its links (`UpdateLinks` 0) and read, computed fully and read; then opened with `Src.xlsx` and read.
   */
  | { external: string; cells: Record<string, string> }
) & { pt?: string }

export interface PivotCase {
  id: string
  /** The rows (`DATA`, `sales` by default) and how the cache takes them (a range by default). */
  data?: string
  source?: SourceKind
  /** The book's date system, 1900 by default. */
  dateSystem?: 1904
  /** A saved book of another set (`basic/sum`) to start from instead of the rows; its pivot table `PT1` is there. */
  from?: string
  /**
   * An xlsx file to start from (a path from the package's folder, or absolute), such as one the engine wrote: Excel
   * opens it as it is, and each phase reads every pivot table of the book, keyed `sheet!name`.
   */
  file?: string
  ops: PivotOp[]
  /** Ranges read at the end of each phase (`G!A1:A20`). */
  watch?: string[]
}

export interface PivotSet {
  buildPivotCases(): PivotCase[]
}

/**
 * A pivot table's area as Excel showed it: `TableRange1`, `TableRange2` (with the page fields), `DataBodyRange` (empty
 * without one), and each cell of `TableRange2` by rows: its value, its text and its number format; in answers recorded
 * since the set `areas`, also `PageRange`, `RowRange` and `ColumnRange` (empty without one).
 */
export interface PivotArea {
  table: string
  area: string
  body: string
  values: CellValue[][]
  text: string[][]
  formats: string[][]
  page?: string
  rows?: string
  columns?: string
}

/** A slicer's items: name, selected, has data. */
export type SlicerItem = [name: string, selected: boolean, hasData: boolean]

/** The end of a phase: Excel's refusals of its operations (`op <index>: <error>`), each pivot table's area or Excel's error reading it, the cells watched, the slicers' items. */
export interface PivotPhase {
  errors?: string[]
  /** What the operations that answer something answered, by `op <index>`: a data field's caption, `RefreshTable`'s answer, each pivot table's for `refreshAll`. */
  results?: Record<string, unknown>
  pivots: Record<string, PivotArea | { error: string }>
  cells?: Record<string, CellValue>
  slicers?: Record<string, SlicerItem[] | { timeline: [start: number | null, end: number | null, level: number] } | { error: string }>
  /** What an `external` operation read: each read's cells, in the order the operation names them. */
  external?: { reads: Record<string, Record<string, CellValue>>; errors?: string[] }
}

export interface PivotAnswer {
  phases: PivotPhase[]
}
