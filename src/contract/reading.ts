/** Reading a session's book: a cell as a session shows it, queries in pages, Find and Replace, the dependents of ranges, a selection's summary, plain-text copies and values, a cell's rule of data validation, where a sheet's cells end, and windows of a sheet. */
import type { CellRef, CellValue, Clock, ErrorCode, JsonValue, RangeRef, Refusal } from './values.js'
import type { CellInput, CellLink, CellNote } from './book.js'
import type { CellStyle, ConditionalAppearance, DataValidation } from './looks.js'
import type { CellDisplay, CellIssueDetail, CellResult } from './results.js'
import type { EditOptions, EditResult } from './edits.js'
import type { CellTextRefusal } from './typing.js'

/**
 * A changed cell, with `spill`, `spillFill` and `spillOf`, and with `display` its `text` and `color`, as in
 * `CellResult`. A formula whose spill range or `spillFill` changed is listed even when its own value did not; the
 * cells it no longer lists are listed when their value changed. With `display`, a cell whose shown text or colour
 * changed is listed too (a new format, a column's width), with its value.
 */
/**
 * A cell returned by `Session.getCells`: like `CellResult`, plus the cell's own `style` (an index into the workbook's
 * `styles`), `note` and `link`, and the `appearance` that its sheet's conditional formats give it (when they set a
 * number format, `text` and `color` display the value in that format).
 */
export interface CellView extends CellResult {
  style?: number
  note?: CellNote
  link?: CellLink
  appearance?: ConditionalAppearance
  /** On the top-left cell of a legacy array formula, its range (`CellInput.array`); `spill` is also that range. */
  array?: string
  /** `numberFormat` as a format code, as `ViewportCell.numberFormat` returns it: a built-in format as its code in the workbook's locale (14 is `dd.mm.yyyy` in ru-RU). */
  numberFormatCode?: string
}

/** A cell field that `readCells` and `findCells` return: `editText` is the text the cell opens with in edit mode, as in a viewport (`ViewportCell.editText`). */
export type CellReadField = 'value' | 'text' | 'formula' | 'editText' | 'style' | 'numberFormat' | 'note' | 'link' | 'issue'
/** A cell as `readCells` and `findCells` return it: its address and the requested fields. */
export interface ReadCell extends CellRef, Partial<CellView> {
  issue?: CellIssueDetail | null
  /** With the field `editText`: the text the cell opens with in edit mode, as in `ViewportCell.editText`; this is what `findCells` searches with `edit`, except for the apostrophe before text that would otherwise be parsed as something else, which the search leaves out. */
  editText?: string
}
/** Options of a query that returns results in pages (`readCells`, `findCells`, `getDependents`): how many results a page holds and where the previous page stopped. */
export interface PageOptions extends EditOptions {
  /** Maximum number of results returned, 1000 by default, at most 10000. */
  maxCells?: number
  /** An opaque continuation of the same query at the same revision; once the session is edited, it is rejected. */
  cursor?: string
  /** Count all matching cells, even after the page is full; false by default. */
  count?: boolean
}
/** Options of `readCells`: the fields to read and which rows and columns. */
export interface ReadCellsOptions extends PageOptions {
  /**
   * value, text and formula by default. Address and sheet are always returned. `text` is the displayed text in a session
   * opened with `display: false` too; `['value', 'formula']` reads without formatting.
   */
  fields?: CellReadField[]
  /**
   * Only the cells in these rows (for example, the rows a host's filter displays), as in `ViewportOptions.rows`: their
   * row numbers on the sheet, from 1, in ascending order, each listed once; the cells of each range that fall in these
   * rows, with `total` counting only those. A row outside a range contributes nothing to that range.
   */
  rows?: number[]
}
/** Options of `findCells`: what to search and how, in addition to those of `readCells`. */
export interface FindCellsOptions extends ReadCellsOptions {
  /**
   * What to search: the text a cell displays (`text`, the default; Excel's Look in: Values), the text it opens with in
   * edit mode (`edit`: a formula in the workbook's locale, a number in its format, text as is, without an apostrophe;
   * Excel's Look in: Formulas), its stored formula (`formula`) or its value (`value`: the same in every locale, a number
   * as JavaScript writes it, `TRUE` or `FALSE`, an error by its name such as `#DIV/0!`). Regular expressions are not
   * supported.
   */
  in?: 'text' | 'edit' | 'formula' | 'value'
  /** Excel's Match case; case-insensitive by default (`straße` matches `STRAẞE`, `i` does not match `İ`). */
  caseSensitive?: boolean
  /** The query is interpreted as Excel's Find interprets it: `*` matches any text, `?` one character (a character outside the BMP takes two), and `~` before a character makes it literal. Without it, the query is literal. */
  wildcards?: boolean
  /** Excel's Match entire cell contents: the entire text must match the query, not just a part of it. */
  wholeCell?: boolean
  /** Excel's Search: By Columns: each range's cells are searched column by column; row by row without it. */
  byColumns?: boolean
  /**
   * Cells in hidden rows and columns are skipped as Excel's Find skips them: always when searching displayed text
   * (`text`, Look in: Values), and in every kind of search on a filtered sheet, where the cells are visited block by
   * block, one block of cells between hidden rows and columns at a time (`docs/live-books.md`). Without it, every cell
   * in the ranges is searched.
   */
  skipHidden?: boolean
  /**
   * The cell after which Find starts, like `Range.Find`'s `After` (a user's Find Next from the active cell): cells are
   * visited in Find's order starting from the one after it, wrapping around to the first cell of the ranges and on to
   * this cell, which comes last; it must be a cell within the ranges.
   */
  after?: CellRef
}

/** Options of `Session.replaceInCells`: Find's options as `findCells` accepts them, except what to search, which is always the text a cell opens with in edit mode (Look in: Formulas), and the search order. */
export interface ReplaceOptions extends EditOptions, Pick<FindCellsOptions, 'caseSensitive' | 'wildcards' | 'wholeCell' | 'skipHidden' | 'rows'> {}

/**
 * What `Session.replaceInCells` did: the edit's result and the number of cells whose text it replaced (`replaced`,
 * counting a legacy array formula as one); the cells whose new text Excel would reject, left unchanged, with the
 * reasons (`refused`).
 */
export interface ReplaceResult extends EditResult {
  replaced: number
  refused?: (Refusal<CellTextRefusal> & CellRef)[]
}

/** A page of a query's results: the revision and clock it read, whether more results remain and where to continue. */
export interface QueryPage {
  revision: number
  clock: Clock
  truncated: boolean
  cursor?: string
  /** With count: the number matched over the complete query, not just this page. */
  total?: number
}
/** What `readCells` and `findCells` return: a page of cells. */
export interface ReadCellsResult extends QueryPage { cells: ReadCell[] }
/** A cell whose formula refers to the ranges passed to `getDependents`, with the references that point to them. */
export interface DependentCell extends CellRef {
  /** Matching source references remain rectangles, never expanded into individual cells. */
  references: RangeRef[]
  referencesTruncated?: boolean
}
/** Options of `getDependents`: the page options and how many references to return for each cell. */
export interface DependentsOptions extends PageOptions {
  /** Maximum matching source references per dependent, 64 by default, at most 1000. */
  maxReferences?: number
}
/** What `getDependents` returns: a page of the cells whose formulas refer to the ranges. */
export interface DependentsResult extends QueryPage { dependents: DependentCell[] }
/** Options of `summarizeRanges`: only the cells in these rows (the rows a host's filter displays), as in `ReadCellsOptions.rows`. */
export interface SummaryOptions extends EditOptions {
  rows?: number[]
}
/** What `summarizeRanges` returns, as Excel's status bar shows it for a selection: how many cells are not empty, how many hold numbers and errors, the sum of the numbers, and the smallest and largest number. */
export interface RangeSummary {
  revision: number
  clock: Clock
  cells: number
  numbers: number
  errors: number
  /** Compensated sum; NUM when it overflows the finite number range. */
  sum: number | { error: 'NUM' }
  min: number | null
  max: number | null
}

/** Options of `Session.copyText`, which copies a selection as plain text. */
export interface CopyTextOptions {
  /** A trailing CRLF, as on Windows; false for Mac. True by default. */
  trailingNewline?: boolean
  /** Maximum number of visible cells to copy, 20,000 by default. A larger copy is rejected, never truncated. */
  maxCells?: number
  /** Maximum length of the encoded text, 16 Mi characters by default, checked before its rows are assembled. */
  maxChars?: number
}

/** Options of `Session.getValues`. */
export interface ValuesOptions {
  /** Maximum number of cells in the rectangle, 1,000,000 by default. A larger one is rejected, never truncated. */
  maxCells?: number
}

/** Options of `Session.getViewport` and `viewportOf`. */
export interface ViewportOptions {
  /** Also return each formula as displayed in the workbook's locale (`ViewportCell.formulaLocal`); false by default. */
  formulaLocal?: boolean
  /**
   * Also return, for each cell with an error and each formula stopped by a circular reference, the reason and the cell
   * where it originated (`ViewportCell.issue`); false by default.
   */
  diagnostics?: boolean
  /** Maximum number of cells a viewport returns, 20,000 by default: for a larger one, the first rows that fit (`Viewport.truncated`). */
  maxCells?: number
  /**
   * Only these rows of `range` (for example, the rows a filter displays): their row numbers on the sheet, from 1, in
   * ascending order, each listed once. The viewport returns them in this order, with only the merged areas and tables
   * that overlap its cells.
   */
  rows?: number[]
  /** Only these columns of `range`, by number (A is 1), in ascending order, each listed once, as with `rows`. */
  columns?: number[]
  /** Also return, for each cell, the text it opens with in edit mode (`ViewportCell.editText`); false by default. */
  editText?: boolean
}

/**
 * A cell of a viewport (`Session.getViewport`): its value as `getResult()` returns it, with the cell's formatting fully
 * resolved: its style and number format, including those inherited from rows and columns, the effect of its
 * conditional formats, its note and its hyperlink.
 */
export interface ViewportCell extends Partial<CellDisplay> {
  value: CellValue
  /** The canonical formula text, as in `CellResult.formula`. */
  formula?: string
  /** With `formulaLocal`: the formula as Excel displays it in the workbook's locale (`Range.FormulaLocal` without `=`), and inside its table the way Excel displays it there (`[@Qty]`). */
  formulaLocal?: string
  spill?: string
  spillFill?: CellValue
  spillOf?: string
  /** On the top-left cell of a legacy array formula, its range (`CellInput.array`). */
  array?: string
  /** As `CellResult.meta`: on the session's thread the workbook's own object, which must not be modified. */
  meta?: JsonValue
  /** The cell's own style, an index into the workbook's `styles` as in `CellView.style`; the style object is in `Viewport.styles`. */
  style?: number
  /**
   * The style the cell uses, an index into the workbook's `styles` (the style object is in `Viewport.styles`): its own,
   * or else its row's, or else its column's. It is always a style in the table: a cell where a styled row crosses a
   * styled column uses the row's, as Excel displays an empty cell from a file there, and a combination of the two would
   * be a style of the cell's own (`Session.patchCells`).
   */
  effectiveStyle?: number
  /** The number format the cell is displayed in, as a format code: its own, or else its row's, or else its column's; a built-in format as its code in the workbook's locale (14 is `dd.mm.yyyy` in ru-RU). */
  numberFormat?: string
  /** Where `numberFormat` comes from. */
  numberFormatSource?: 'cell' | 'row' | 'column'
  /** The effect of the sheet's conditional formats on the cell, given the current values; when they set a number format, `text` and `color` display the value in that format. */
  appearance?: ConditionalAppearance
  note?: CellNote
  link?: CellLink
  /** With `diagnostics`, on a cell with an error or a formula stopped by a circular reference: as returned by `Session.getCellIssue`. */
  issue?: CellIssueDetail
  /**
   * With `editText`: the text the cell opens with in edit mode, as Excel's formula bar displays it: a formula with `=`
   * in the workbook's locale (`formulaLocal`); in every cell of a legacy array formula's range, its formula (without
   * the braces the bar shows around it); a value as `entryText` writes it in the cell's number format; and empty for a
   * cell without a value or formula of its own (an empty cell, or one that a dynamic array spilled into).
   */
  editText?: string
}

/** A column of a viewport: its number (A is 1) and its width as in `ColumnInput.width`, or Excel's default where it has none. */
export interface ViewportColumn {
  column: number
  width: number
  hidden?: true
}

/** A row of a viewport: its number on the sheet, from 1, and its height in points, or Excel's default (15) where it has none. */
export interface ViewportRow {
  row: number
  height: number
  hidden?: true
}

/** A table that overlaps a viewport (`Viewport.tables`): its name, its range, and whether it has header and totals rows. */
export interface ViewportTable {
  name: string
  /** The entire table, including its header and totals rows (`TableInput.range`). */
  range: string
  /** Its first row is a header row (`TableInput.headers`, true by default). */
  headers: boolean
  /** Its last row is a totals row (`TableInput.totals`, false by default). */
  totals: boolean
}

/** What a grid needs to draw a viewport of a sheet (`Session.getViewport`). */
export interface Viewport {
  /**
   * The rectangle returned (`B2:D9`): the requested one, or, when it holds more than `maxCells` cells, the part up to
   * its first rows that fit (of `ViewportOptions.rows`, when given).
   */
  range: string
  /** The requested viewport held more than `maxCells` cells: `range` is the part returned. */
  truncated: boolean
  /** By address, row by row: every cell with something to display, including an empty cell with its row's or column's style or number format, a conditional format's appearance, a note or a hyperlink. */
  cells: Record<string, ViewportCell>
  /** The styles the viewport's cells have or use (`ViewportCell.style`, `effectiveStyle`), by their index in the workbook's `styles`: copies owned by the result. */
  styles: Record<number, CellStyle>
  /** The columns of `range` from left to right, or those of `ViewportOptions.columns`. */
  columns: ViewportColumn[]
  /** The rows of `range` from top to bottom, or those of `ViewportOptions.rows`. */
  rows: ViewportRow[]
  /** The merged areas that overlap the viewport's cells, in full. */
  merges: string[]
  /** The tables that overlap the viewport's cells, in full, in the sheet's order. */
  tables: ViewportTable[]
  /** The sheet's frozen panes (`SheetInput.frozen`): its first `rows` rows and `columns` columns stay visible, 0 for none. */
  frozen: { rows: number; columns: number }
  /** Where the sheet's cells end, as returned by `Session.getUsedRange`. */
  usedRange: SheetExtent
  /** The session revision that the viewport reflects (`Session.revision`): a host applies the changes of later edits on top of it (`EditResult.revision`). */
  revision: number
}

/** Options of `viewportOf`: those of `getViewport`, and `appearance`. */
export interface SavedViewportOptions extends ViewportOptions {
  /**
   * Also return, for each cell, the `appearance` its sheet's conditional formats give it, and its `text` and `color` in a
   * number format they set, as `getViewport` returns them in a session of the workbook on the result's clock; false by
   * default. A result holds no appearance: the rules are calculated over the result's values, but its formulas are not
   * recalculated (TODAY in a rule is the result's day). The first viewport with this option over a workbook and a result
   * reads all their values; later ones cost only what the viewport and its rules read.
   */
  appearance?: boolean
}

/**
 * A cell of a viewport over a saved result (`viewportOf`): like `ViewportCell`, with its value, `text` and `color` as
 * the result has them; without `text` and `color` when the result was calculated without display text. The
 * `appearance` of conditional formats is present only with `SavedViewportOptions.appearance`: without it, the cell's
 * `text` displays the value in the cell's own format, not a rule's.
 */
export type SavedViewportCell = ViewportCell

/** A viewport over a saved result (`viewportOf`): like `Viewport`, but without a session revision, and with `SavedViewportCell` cells. */
export interface SavedViewport extends Omit<Viewport, 'revision' | 'cells'> {
  cells: Record<string, SavedViewportCell>
}

/**
 * A cell's data validation (`Session.getValidation`): the rule that covers it, the rule's formulas as they apply at the
 * cell (relative references shifted from the rule's first cell), and, for a list, the items its dropdown offers.
 */
export interface CellValidation {
  rule: DataValidation
  formula1?: string
  formula2?: string
  /**
   * A list's items in order, calculated in the workbook's current state: a literal list's items as written (`text`, with
   * `value` the same text), or the cells of the reference its formula returns, each with its value and the text it
   * displays in its number format (blank cells after the last value are left out). An error when the formula returns
   * neither a reference nor a list. `input` is what Excel writes into the cell when the item is picked from the
   * dropdown, for `setCells`, which keeps the cell's format: a cell's value as is (`001` stays text, a number displayed
   * as `50%` is 0.5), a literal item parsed as typed in en-US, the locale in which a file stores the list (`1/2` is the
   * 2nd of January, `d-mmm`), or null for a blank cell. Excel checks no rule on a picked item.
   */
  list?: { value: CellValue; text: string; input: CellInput | null }[] | { error: ErrorCode }
}

/** Where a sheet's cells end (`Session.getUsedRange`): all content lies within its first `rows` rows and `columns` columns; 0 and 0 for an empty sheet. */
export interface SheetExtent {
  rows: number
  columns: number
}

/** The direction in which Control plus an arrow key moves from a cell (`Session.getEnd`). */
export type EndDirection = 'up' | 'down' | 'left' | 'right'
