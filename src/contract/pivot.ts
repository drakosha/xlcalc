/**
 * PivotTables, their caches, slicers and timelines, as a workbook holds them (`BookInput.pivotCaches`,
 * `SheetInput.pivotTables`): plain data, no XML, so a workbook with them goes to a pool's threads as any other.
 */
import type { NumberFormat } from './book.js'
import type { CellStyle } from './looks.js'
import type { ErrorCode } from './values.js'

/**
 * A value of a pivot cache: a blank (`null`, Excel's missing item), a number, text, a boolean, an error, or a date and
 * time as Excel stores it in a cache, ISO 8601 without a zone (`{ date: '2025-01-31T00:00:00' }`).
 */
export type PivotValue = null | number | string | boolean | { error: ErrorCode } | { date: string }

/** The functions that summarize a value field or make a field's subtotals (Excel's Summarize Values By). */
export type PivotFunction = 'sum' | 'count' | 'average' | 'max' | 'min' | 'product' | 'countNums' | 'stdDev' | 'stdDevp' | 'var' | 'varp'

/**
 * Where a pivot cache takes its data from. `range`: cells of a sheet; `table`: a table by name; `name`: a defined name;
 * each in the workbook itself, or in another one (`book`, the link's target as the file stores it). The other kinds are
 * held and not refreshed: several ranges consolidated, an external connection (`connection`, its id in the file's
 * connections), the workbook's data model, scenarios.
 */
export type PivotSource =
  | { type: 'range'; sheet: string; range: string; book?: string }
  | { type: 'table'; table: string; book?: string }
  | { type: 'name'; name: string; book?: string }
  | { type: 'consolidation'; ranges: { sheet?: string; range: string; book?: string }[] }
  | { type: 'external' | 'dataModel'; connection?: number }
  | { type: 'scenario' }

/** A pivot cache: the source's fields and, when it saved them, its rows, which the workbook's PivotTables summarize. */
export interface PivotCache {
  /** The cache's id in the workbook (xlsx `cacheId`), by which PivotTables, slicers and timelines name it; unique. */
  id: number
  source: PivotSource
  /** The source's columns, then the fields grouping adds and the calculated fields, in the cache's order. */
  fields: PivotCacheField[]
  /**
   * The source's rows as the cache saved them (xlsx `pivotCacheRecords`), a value per field of the source: for a field
   * with `items`, the index of its item; otherwise the value itself. Absent when the cache saved no rows.
   */
  records?: PivotValue[][]
  /** Calculated items of its fields. */
  calculatedItems?: PivotCalculatedItem[]
  /** Refresh the PivotTables when the file is opened (Refresh data when opening the file). */
  refreshOnLoad?: boolean
  /** The rows are saved with the file (Save source data with file, true by default). */
  saveData?: boolean
  /** Items no longer in the source each field retains (Number of items to retain per field): 0 for none, absent for automatic. */
  retainItems?: number
  /** Refresh is allowed (true by default). */
  enableRefresh?: boolean
  /**
   * A PivotTable of the cache was changed by a session without `xlcalc/pivot`, so its cells are not laid out again: the
   * file is written with Refresh data when opening the file, and a refresh clears it.
   */
  stale?: boolean
}

/** The kinds of value a cache field holds, which Excel keeps beside its items (xlsx `sharedItems` flags). */
export interface PivotFieldFacts {
  /** The kinds present; a field of no rows has none. */
  types: ('blank' | 'number' | 'string' | 'boolean' | 'error' | 'date')[]
  /** Every number is a whole number. */
  integer?: boolean
  /** Some text is longer than 255 characters. */
  longText?: boolean
  /** The least and greatest number, or the earliest and latest date, of a field without its items listed. */
  min?: number | { date: string }
  max?: number | { date: string }
}

/** A field of a pivot cache: a column of the source, a field grouping made, or a calculated field. */
export interface PivotCacheField {
  name: string
  /** The source column's number format (Excel's field Number Format). */
  numberFormat?: NumberFormat
  /** The field's distinct values (xlsx `sharedItems`), in the cache's order; absent when the cache does not list them (numbers, dates). */
  items?: PivotValue[]
  /** Indexes of `items` no longer in the source, kept by `retainItems`. */
  unused?: number[]
  facts?: PivotFieldFacts
  /** A calculated field's formula over other fields by name (`Amount*2`, `'Unit Price'*Qty`), as the file stores it. */
  formula?: string
  group?: PivotFieldGroup
}

/** How a date is grouped, or numbers by intervals (`numbers`). */
export type PivotGroupBy = 'numbers' | 'seconds' | 'minutes' | 'hours' | 'days' | 'months' | 'quarters' | 'years'

/**
 * A field's grouping (xlsx `fieldGroup`). Dates by several levels make a field per level after the first, each with the
 * date field as its `base` and the next coarser level as its `parent`; a manual group makes a field over the grouped
 * field's items (`discrete`).
 */
export interface PivotFieldGroup {
  /** The index of the field whose values are grouped; absent when it is this field. */
  base?: number
  /** The index of the field that groups this one further. */
  parent?: number
  /** Dates by a unit or numbers by intervals, between `start` and `end`, each automatic unless set (Starting at, Ending at, By). */
  range?: {
    by: PivotGroupBy
    start?: number | { date: string }
    end?: number | { date: string }
    autoStart?: boolean
    autoEnd?: boolean
    /** The interval: numbers per group, or days per group when `by` is `days`. */
    interval?: number
  }
  /** Manual groups: for each item of the base field, the index of its group in `items`. */
  discrete?: number[]
  /** The groups' labels (`<1/1/2024`, `Qtr1`, `Jan`, `0-9`, `Group1`). */
  items?: PivotValue[]
}

/** A field on an axis: the index of a cache field, or `values`, the place of the value fields when there are several (Σ Values). */
export type PivotAxisField = number | 'values'

/** A field on the Filters area (xlsx `pageField`): the field, and the item shown, every item without it. */
export interface PivotPageField {
  field: number
  /** The index into the field's `items` of the one item shown. */
  item?: number
}

/** A field on the Values area (xlsx `dataField`). */
export interface PivotDataField {
  /** The index of the cache field summarized. */
  field: number
  /** Its name in the PivotTable (`Sum of Amount`, Custom Name), unique among the PivotTable's value fields. */
  name?: string
  /** `sum` by default. */
  function?: PivotFunction
  /** Show Values As; `normal` by default. */
  showAs?: PivotShowAs
  /** The base field of `difference`, `percent`, `percentDiff`, `runTotal`, `percentOfRunningTotal`, `percentOfParent`, ranks and `index`. */
  baseField?: number
  /** The base item: an index into the base field's `items`, or the previous or next item. */
  baseItem?: number | 'previous' | 'next'
  numberFormat?: NumberFormat
}

/** Show Values As, Excel's 15 calculations. */
export type PivotShowAs =
  | 'normal'
  | 'percentOfTotal'
  | 'percentOfColumn'
  | 'percentOfRow'
  | 'percent'
  | 'percentOfParentRow'
  | 'percentOfParentColumn'
  | 'percentOfParent'
  | 'difference'
  | 'percentDiff'
  | 'runTotal'
  | 'percentOfRunningTotal'
  | 'rankAscending'
  | 'rankDescending'
  | 'index'

/** A PivotTable's form: compact, outline, or tabular (Report Layout). */
export type PivotLayout = 'compact' | 'outline' | 'tabular'

/** An item of a PivotTable's field, in the field's order on the PivotTable. */
export interface PivotFieldItem {
  /** The index into the cache field's `items` (or its group's), absent for a calculated item that the cache does not list. */
  item?: number
  /** Its caption on the PivotTable, in place of the value. */
  caption?: string
  /** Filtered out (unchecked). */
  hidden?: boolean
  /** Its details are collapsed (shown by default). */
  collapsed?: boolean
  /** A calculated item, its formula in the cache's `calculatedItems`. */
  calculated?: boolean
  /** No longer in the source. */
  missing?: boolean
}

/** A field of a PivotTable, one per cache field: how it shows on its axis, whichever that is. */
export interface PivotField {
  /** Its custom name on the PivotTable. */
  name?: string
  /** Its items in the PivotTable's order; absent for a field on no axis, or on Values alone. */
  items?: PivotFieldItem[]
  /** The functions of its subtotals: absent for automatic (the value field's own), `[]` for none. */
  subtotals?: PivotFunction[]
  /** Subtotals above the items' group (true by default), in outline and compact forms. */
  subtotalTop?: boolean
  /** The subtotal's label, in place of `North Total`. */
  subtotalCaption?: string
  /** Its form, the PivotTable's without it. */
  layout?: PivotLayout
  /** Repeat item labels. */
  repeatLabels?: boolean
  /** A blank line after each item. */
  blankLine?: boolean
  /** A page break after each item. */
  pageBreak?: boolean
  /** Show items with no data. */
  showEmpty?: boolean
  /** Sort its items by label, or by a value field's values (`by`), ascending or descending; the cache's order without it. */
  sort?: { order: 'ascending' | 'descending'; by?: PivotSortBy }
  /** Sort automatically every time the report is updated (true by default). */
  autoSort?: boolean
  /** Include new items in a manual filter. */
  includeNewItems?: boolean
  /** Several items may be shown at once on the Filters area (Select Multiple Items). */
  multipleItems?: boolean
  numberFormat?: NumberFormat
}

/** Sorting by a value field: its index in `PivotTable.values`, at the items of the other axis given (a column's values), at the grand total without them. */
export interface PivotSortBy {
  value: number
  at?: { field: number; item: number }[]
}

/** The kinds of a PivotTable's filter: by label, by value, Top 10, by date. */
export type PivotFilterType =
  | 'captionEqual' | 'captionNotEqual' | 'captionBeginsWith' | 'captionNotBeginsWith' | 'captionEndsWith' | 'captionNotEndsWith'
  | 'captionContains' | 'captionNotContains' | 'captionGreaterThan' | 'captionGreaterThanOrEqual' | 'captionLessThan'
  | 'captionLessThanOrEqual' | 'captionBetween' | 'captionNotBetween'
  | 'valueEqual' | 'valueNotEqual' | 'valueGreaterThan' | 'valueGreaterThanOrEqual' | 'valueLessThan' | 'valueLessThanOrEqual'
  | 'valueBetween' | 'valueNotBetween'
  | 'count' | 'percent' | 'sum'
  | 'dateEqual' | 'dateNotEqual' | 'dateOlderThan' | 'dateOlderThanOrEqual' | 'dateNewerThan' | 'dateNewerThanOrEqual'
  | 'dateBetween' | 'dateNotBetween' | 'tomorrow' | 'today' | 'yesterday' | 'nextWeek' | 'thisWeek' | 'lastWeek'
  | 'nextMonth' | 'thisMonth' | 'lastMonth' | 'nextQuarter' | 'thisQuarter' | 'lastQuarter' | 'nextYear' | 'thisYear'
  | 'lastYear' | 'yearToDate' | 'Q1' | 'Q2' | 'Q3' | 'Q4' | 'M1' | 'M2' | 'M3' | 'M4' | 'M5' | 'M6' | 'M7' | 'M8' | 'M9'
  | 'M10' | 'M11' | 'M12'

/**
 * A filter of a field by label, value, Top 10 or date. `args` are the dialog's one or two values: text for a label
 * filter, numbers for a value filter, ISO dates for a date filter, the count, percent or sum for Top 10.
 */
export interface PivotFilter {
  field: number
  type: PivotFilterType
  args?: (string | number)[]
  /** The value field a value filter or Top 10 tests: its index in `PivotTable.values`. */
  value?: number
  /** Top 10's Bottom (top by default). */
  bottom?: boolean
  /** A date filter's whole days (false: the dates' times count). */
  wholeDays?: boolean
  /** The filter's order among the field's filters. */
  order?: number
  name?: string
  description?: string
}

/**
 * A line of the PivotTable's saved layout (xlsx `rowItems` and `colItems`): a row on rows, a column on columns. `items`
 * are, for each field of the axis from `repeat` on, the index into that field's `items` (or into the value fields, for
 * `values`); the fields before `repeat` are the line before's.
 */
export interface PivotLine {
  /** A data line by default; a field's subtotal (automatic, or of a function), the grand total, a blank line. */
  type?: 'subtotal' | PivotFunction | 'grand' | 'blank'
  /** Fields at the start taken from the line before. */
  repeat?: number
  items: number[]
  /** The value field of the line, its index in `PivotTable.values`, when the value fields lie across this axis. */
  value?: number
}

/**
 * A PivotTable's place: `range` the table without its Filters area, which takes the rows `pageRows` (and the columns
 * `pageColumns`, wrapped) above it with a blank row between; the first row of the table's headers, the first row and
 * column of its values, from the range's top left at 0.
 */
export interface PivotLocation {
  range: string
  firstHeaderRow: number
  firstDataRow: number
  firstDataColumn: number
  pageRows?: number
  pageColumns?: number
}

/** A PivotTable's look, as Excel's PivotTable Style Options. */
export interface PivotStyle {
  /** `PivotStyleLight16` and the others; none without it. */
  name?: string
  rowHeaders?: boolean
  columnHeaders?: boolean
  rowStripes?: boolean
  columnStripes?: boolean
  lastColumn?: boolean
}

/** A PivotTable on a sheet (`SheetInput.pivotTables`). */
export interface PivotTable {
  /** Unique among the sheet's PivotTables. */
  name: string
  /** The `id` of the workbook's pivot cache it summarizes. */
  cache: number
  location: PivotLocation
  /** One per field of the cache, in its order. */
  fields: PivotField[]
  rows?: PivotAxisField[]
  columns?: PivotAxisField[]
  pages?: PivotPageField[]
  values?: PivotDataField[]
  /** The value fields lie across the rows, one line each (true) or across the columns. */
  valuesOnRows?: boolean
  /** The place of the value fields among the axis's fields when there is one, last without it. */
  valuesPosition?: number
  filters?: PivotFilter[]
  /** Grand totals of the rows (a column at the right) and of the columns (a row at the bottom), each shown by default. */
  grandTotals?: { rows?: boolean; columns?: boolean }
  /** Its form, compact by default; a field can have its own. */
  layout?: PivotLayout
  /** Characters of indent per level in compact form (1 by default). */
  indent?: number
  /** Repeat item labels of every field without its own (Repeat All Item Labels). */
  repeatLabels?: boolean
  /** Merge and center cells with labels. */
  mergeLabels?: boolean
  /** Field headers shown (true by default). */
  fieldHeaders?: boolean
  /** Its labels, each Excel's in the workbook's language without it. */
  captions?: PivotCaptions
  /** Show for errors and empty cells: the text shown in place of an error, and in place of an empty value (shown by default, as nothing). */
  errors?: { show: boolean; text?: string }
  empty?: { show: boolean; text?: string }
  /** Report filter fields: across then down (down then across by default), and fields per column or row (0, unlimited). */
  pagesAcrossFirst?: boolean
  pagesWrap?: number
  /** Subtotal filtered page items; mark totals with `*`. */
  subtotalHiddenItems?: boolean
  asteriskTotals?: boolean
  /** Use custom lists when sorting (true by default). */
  customListSort?: boolean
  /** Allow multiple filters per field. */
  multipleFieldFilters?: boolean
  /** Show the expand and collapse buttons (true by default). */
  showDrill?: boolean
  /** Autofit column widths on update (true by default) and preserve cell formatting (true by default). */
  autofitColumns?: boolean
  preserveFormatting?: boolean
  /** Classic PivotTable layout (dragging fields in the grid). */
  classicLayout?: boolean
  style?: PivotStyle
  /** Enable show details (true by default). */
  enableDrill?: boolean
  /** Alt Text: its title and description. */
  altText?: { title?: string; description?: string }
  /** Formats of its areas, which Excel applies after each update (with `preserveFormatting`); in its order, a later one over an earlier. */
  formats?: PivotFormat[]
  /** The saved layout: the lines of the rows and of the columns as Excel last laid them out. */
  rowLines?: PivotLine[]
  columnLines?: PivotLine[]
}

/** A PivotTable's labels (`Grand Total`, `Row Labels`, `Column Labels`, `Values`). */
export interface PivotCaptions {
  grandTotal?: string
  rowHeader?: string
  columnHeader?: string
  values?: string
}

/** A calculated item of a cache's field (xlsx `calculatedItem`), on every PivotTable of the cache. */
export interface PivotCalculatedItem {
  field: number
  /** The index of the item in the field's `items`. */
  item: number
  /** Its formula over the field's items (`North+South`), as the file stores it. */
  formula: string
}

/** A PivotTable by its sheet and name, as a slicer or a timeline names it. */
export interface PivotRef {
  sheet: string
  name: string
}

/** A place on a sheet by a drawing's anchor: the cells of its corners, with offsets in EMU (914,400 to an inch). */
export interface DrawingPlace {
  from: string
  to: string
  fromOffset?: [x: number, y: number]
  toOffset?: [x: number, y: number]
}

/** A slicer's cache (xlsx x14 `slicerCacheDefinition`): its field and the items chosen, shared by its slicers. */
export interface SlicerCache {
  /** Its name in the workbook (`Slicer_Region`), unique. */
  name: string
  /** The field's name in the source. */
  field: string
  /** The PivotTables it filters, all of one cache. */
  pivots: PivotRef[]
  /** The pivot cache's `id`; absent for a table's slicer. */
  cache?: number
  /** A table's slicer: the table and its column, in place of a pivot cache. */
  table?: { name: string; column: string }
  /** The field's items with whether each is chosen (chosen by default) and has data. */
  items?: { item: number; selected?: boolean; noData?: boolean }[]
  /** Item sorting, ascending by default; with custom lists (true by default). */
  sort?: 'ascending' | 'descending'
  customListSort?: boolean
  /** Items with no data: hidden (false by default), marked (true by default), shown last (true by default). */
  noData?: { hide?: boolean; mark?: boolean; last?: boolean }
  /** Items deleted from the source shown (true by default). */
  showMissing?: boolean
}

/** A slicer on a sheet (`SheetInput.slicers`). */
export interface Slicer {
  /** Its name in the workbook, unique. */
  name: string
  /** The name of its `SlicerCache`. */
  cache: string
  caption?: string
  /** Its header shown (true by default). */
  showCaption?: boolean
  place?: DrawingPlace
  /** `SlicerStyleLight1` and the others. */
  style?: string
  columns?: number
  /** The height of its buttons, in EMU. */
  rowHeight?: number
  /** The first item shown, by its index. */
  startItem?: number
  locked?: boolean
}

/** A timeline's level: years, quarters, months, days. */
export type TimelineLevel = 'years' | 'quarters' | 'months' | 'days'

/** A timeline's cache (xlsx x15 `timelineCacheDefinition`): its date field and the dates chosen. */
export interface TimelineCache {
  name: string
  /** The date field's name in the source. */
  field: string
  pivots: PivotRef[]
  cache: number
  /** The dates chosen, ISO 8601 without a zone, all without it. */
  selection?: { start: string; end: string }
  /** The dates the timeline spans. */
  bounds?: { start: string; end: string }
}

/** A timeline on a sheet (`SheetInput.timelines`). */
export interface Timeline {
  name: string
  /** The name of its `TimelineCache`. */
  cache: string
  caption?: string
  place?: DrawingPlace
  /** The level shown, months by default, and the level the selection was made at. */
  level?: TimelineLevel
  selectionLevel?: TimelineLevel
  /** Its header, the selection's label, the level, the scrollbar; each but the scrollbar shown by default. */
  showHeader?: boolean
  showSelectionLabel?: boolean
  showLevel?: boolean
  showScrollbar?: boolean
  /** The date scrolled to, ISO 8601. */
  scrollPosition?: string
  /** `TimeSlicerStyleLight1` and the others. */
  style?: string
}

/**
 * A part of a PivotTable (xlsx `pivotArea`), which formats and conditional formats apply to: by default the values and
 * labels at the items its references name, on whatever lines the PivotTable lays them out.
 */
export interface PivotArea {
  /** `normal` by default: the cells its references and flags pick; `data` the values, `all` the whole PivotTable, `origin` the top left cell, `button` a field's button, `topEnd` and `topRight` the cells right of the column fields, `none` nothing. */
  type?: 'normal' | 'none' | 'data' | 'all' | 'origin' | 'button' | 'topEnd' | 'topRight'
  /** The field whose labels or button it is. */
  field?: PivotAxisField
  /** The axis and position of a button or of a field's labels. */
  axis?: 'rows' | 'columns' | 'pages' | 'values'
  fieldPosition?: number
  /** Values only (true by default), or labels only. */
  dataOnly?: boolean
  labelOnly?: boolean
  /** The grand total's row or column. */
  grandRow?: boolean
  grandColumn?: boolean
  /** In outline form (true by default). */
  outline?: boolean
  /** A part of the area's cells, relative to its top left (`A1:B1`). */
  offset?: string
  /** Items by their index in the cache field's items rather than the PivotTable field's. */
  cacheIndex?: boolean
  collapsedLevelsAreSubtotals?: boolean
  references?: PivotAreaReference[]
}

/** A field of a `PivotArea` and the items picked on it, all without them. */
export interface PivotAreaReference {
  field: PivotAxisField
  /** Indexes into the field's `items` (or the value fields, for `values`). */
  items?: number[]
  /** The items are picked (true by default), by position among those shown, relative to the cell. */
  selected?: boolean
  byPosition?: boolean
  relative?: boolean
  /** The field's subtotals picked: automatic, or of each function. */
  subtotals?: ('subtotal' | PivotFunction)[]
}

/** A format of a PivotTable's area (xlsx `format`): a style over its cells, each field set replacing the cell's own, or the area cleared (`blank`). */
export interface PivotFormat {
  area: PivotArea
  style?: CellStyle
  numberFormat?: NumberFormat
  blank?: boolean
}

/**
 * A conditional format of a PivotTable (`ConditionalBase.pivot`): the rule stays among the sheet's rules, and this names
 * the PivotTable on its sheet and the areas it applies to, which follow the PivotTable as it is laid out again. `scope`:
 * the selected cells (`selection`, by default), every value of the value field (`data`), or of the value field at the
 * fields of the area (`field`); `type`: how a color scale, data bar or icon set ranks, over all cells, within a row or
 * within a column (`none` by default).
 */
export interface PivotRuleScope {
  pivot: string
  scope?: 'selection' | 'data' | 'field'
  type?: 'none' | 'all' | 'row' | 'column'
  areas: PivotArea[]
}

/**
 * The fields of a PivotTable an edit sets (`editPivotTable`, `addPivotTable`), in the form `PivotTable` holds them: a
 * value replaces the field, null removes it, an omitted field is left as it is; `fields` edits the PivotTable's fields,
 * each by the index of its cache field.
 */
export type PivotTablePatch = { [K in Exclude<keyof PivotTable, 'name' | 'cache' | 'location' | 'fields' | 'rowLines' | 'columnLines'>]?: PivotTable[K] | null } & {
  fields?: PivotFieldPatch[]
}

/** An edit of a PivotTable's field (`PivotTablePatch.fields`): the index of its cache field, and its fields set as in `PivotTablePatch`. */
export type PivotFieldPatch = { field: number } & { [K in keyof PivotField]?: PivotField[K] | null }

/** The fields of a slicer an edit sets (`editSlicer`), as in `PivotTablePatch`. */
export type SlicerPatch = { [K in Exclude<keyof Slicer, 'name' | 'cache'>]?: Slicer[K] | null }

/** The fields of a slicer's cache an edit sets (`editSlicer`): the PivotTables it filters, the items chosen, their order and display. */
export type SlicerCachePatch = { [K in 'pivots' | 'items' | 'sort' | 'customListSort' | 'noData' | 'showMissing']?: SlicerCache[K] | null }

/** The fields of a timeline an edit sets (`editTimeline`), as in `PivotTablePatch`. */
export type TimelinePatch = { [K in Exclude<keyof Timeline, 'name' | 'cache'>]?: Timeline[K] | null }

/** The fields of a timeline's cache an edit sets (`editTimeline`): the PivotTables it filters and the dates chosen. */
export type TimelineCachePatch = { [K in 'pivots' | 'selection']?: TimelineCache[K] | null }
