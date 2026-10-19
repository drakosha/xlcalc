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
 * A cell of `Session.getCells`: as `CellResult`, and the cell's own `style` (an index of the book's `styles`), `note` and
 * `link`, and the `appearance` its sheet's conditional formats give it (with a number format of theirs, `text` and
 * `color` show the value by it).
 */
export interface CellView extends CellResult {
  style?: number
  note?: CellNote
  link?: CellLink
  appearance?: ConditionalAppearance
  /** On the top-left cell of an array formula of the old kind, its area (`CellInput.array`); `spill` is that area too. */
  array?: string
  /** `numberFormat` as a code, as `ViewportCell.numberFormat` gives one: a built-in format by the code of the book's locale (14 is `dd.mm.yyyy` in ru-RU). */
  numberFormatCode?: string
}

/** A field `readCells` and `findCells` give of a cell: `editText` is the text it opens with for editing, as a window's (`ViewportCell.editText`). */
export type CellReadField = 'value' | 'text' | 'formula' | 'editText' | 'style' | 'numberFormat' | 'note' | 'link' | 'issue'
/** A cell as `readCells` and `findCells` give it: its address and the fields asked for. */
export interface ReadCell extends CellRef, Partial<CellView> {
  issue?: CellIssueDetail | null
  /** With the field `editText`: the text the cell opens with for editing, as `ViewportCell.editText`; what `findCells` reads in `edit`, but for the apostrophe before a text that reads as something else, which it leaves out. */
  editText?: string
}
/** Options of a query that answers in pages (`readCells`, `findCells`, `getDependents`): how many results a page holds and where the last page stopped. */
export interface PageOptions extends EditOptions {
  /** Maximum results returned, 1000 by default, at most 10000. */
  maxCells?: number
  /** Opaque continuation of the same query and revision; an edited session refuses it. */
  cursor?: string
  /** Count all matching cells, even after the page is full; false by default. */
  count?: boolean
}
/** Options of `readCells`: the fields to read and which rows and columns. */
export interface ReadCellsOptions extends PageOptions {
  /** value, text and formula by default. Address and sheet are always returned. */
  fields?: CellReadField[]
  /**
   * Only the cells of these rows (the rows a host's filter shows, say), as `ViewportOptions.rows`: their numbers on the
   * sheet, from 1, in ascending order, each once; the cells of each range in them, `total` counting those alone. A row
   * outside a range leaves that range nothing of it.
   */
  rows?: number[]
}
/** Options of `findCells`: what is searched and how, beside those of `readCells`. */
export interface FindCellsOptions extends ReadCellsOptions {
  /**
   * What is searched: the text a cell shows (`text`, the default; Excel's Look in Values), the text it opens with for
   * editing (`edit`: a formula in the book's locale, a number by its format, a text as it is, without an apostrophe;
   * Excel's Look in Formulas), its stored formula (`formula`) or its value (`value`). No regular expressions.
   */
  in?: 'text' | 'edit' | 'formula' | 'value'
  /** Excel's Match case; any case by default (`straße` is `STRAẞE`, `i` is not `İ`). */
  caseSensitive?: boolean
  /** The query as Excel's Find reads it: `*` any text, `?` one character (two for one outside the BMP), `~` before a character takes it as it is. Literal without it. */
  wildcards?: boolean
  /** Excel's Match entire cell contents: the whole text matches the query, not a part of it. */
  wholeCell?: boolean
  /** Excel's Search By Columns: each range's cells column by column; row by row without it. */
  byColumns?: boolean
  /**
   * Cells of hidden rows and columns are skipped as Excel's Find skips them: always where it reads what cells show
   * (`text`, Look in Values), and in every kind of search on a sheet with a filter, where the cells come by the blocks
   * the hidden lines leave, each block in turn (`docs/live-books.md`). Without it every cell of the ranges is searched.
   */
  skipHidden?: boolean
  /**
   * The cell Find begins after, as `Range.Find`'s `After` (a person's Find Next from the active cell): the cells come in
   * Find's order from the first after it, round to the ranges' first and on to it, itself last; a cell of the ranges.
   */
  after?: CellRef
}

/** `Session.replaceInCells`' options: Find's, as `findCells` takes them, but what it reads, which is the text a cell opens with (Look in Formulas), and the order. */
export interface ReplaceOptions extends EditOptions, Pick<FindCellsOptions, 'caseSensitive' | 'wildcards' | 'wholeCell' | 'skipHidden' | 'rows'> {}

/**
 * What `Session.replaceInCells` did: the edit's result, and how many cells' texts it wrote (`replaced`, an array of the
 * old kind as one); the cells whose new text Excel would not take, left as they were, with why (`refused`).
 */
export interface ReplaceResult extends EditResult {
  replaced: number
  refused?: (Refusal<CellTextRefusal> & CellRef)[]
}

/** A page of a query's answer: the revision and clock it read, whether more is left and where to go on. */
export interface QueryPage {
  revision: number
  clock: Clock
  truncated: boolean
  cursor?: string
  /** With count: the number matched over the complete query, not just this page. */
  total?: number
}
/** What `readCells` and `findCells` answer: a page of cells. */
export interface ReadCellsResult extends QueryPage { cells: ReadCell[] }
/** A cell whose formula reads the ranges `getDependents` was asked about, with the references that read them. */
export interface DependentCell extends CellRef {
  /** Matching source references remain rectangles, never expanded into individual cells. */
  references: RangeRef[]
  referencesTruncated?: boolean
}
/** Options of `getDependents`: those of a page and how many references to give for each cell. */
export interface DependentsOptions extends PageOptions {
  /** Maximum matching source references per dependent, 64 by default, at most 1000. */
  maxReferences?: number
}
/** What `getDependents` answers: a page of the cells whose formulas read the ranges. */
export interface DependentsResult extends QueryPage { dependents: DependentCell[] }
/** `summarizeRanges`' options: only the cells of these rows (the rows a host's filter shows), as `ReadCellsOptions.rows`. */
export interface SummaryOptions extends EditOptions {
  rows?: number[]
}
/** What `summarizeRanges` answers, as Excel's status bar shows a selection: how many cells hold something, how many numbers and errors, their sum, the least and the greatest number. */
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

/** Options of `Session.copyText`, the selection's plain-text copy. */
export interface CopyTextOptions {
  /** A terminal CRLF, as on Windows; false for Mac. True by default. */
  trailingNewline?: boolean
  /** Maximum visible cells to copy, 20,000 by default. A larger copy is refused, never truncated. */
  maxCells?: number
  /** Maximum characters of the encoded text, 16 Mi characters by default, checked before assembling its rows. */
  maxChars?: number
}

/** Options of `Session.getValues`. */
export interface ValuesOptions {
  /** The most cells the rectangle may have, 1,000,000 by default. A larger one is refused, never cut. */
  maxCells?: number
}

/** `Session.getViewport`'s and `viewportOf`'s options. */
export interface ViewportOptions {
  /** Each formula also as the book's locale shows it (`ViewportCell.formulaLocal`); false by default. */
  formulaLocal?: boolean
  /**
   * Each cell with an error, and each formula a circular reference stops, also with its reason and the cell it began in
   * (`ViewportCell.issue`); false by default.
   */
  diagnostics?: boolean
  /** The most cells a window gives, 20,000 by default: of a larger one, its first rows that fit (`Viewport.truncated`). */
  maxCells?: number
  /**
   * Only these rows of `range` (the rows a filter shows, say): their numbers on the sheet, from 1, in ascending order,
   * each once. The window gives them in this order, and only the merged areas and tables that meet its cells.
   */
  rows?: number[]
  /** Only these columns of `range`, by number (A is 1), in ascending order, each once, as `rows`. */
  columns?: number[]
  /** Each cell also with the text it opens with for editing (`ViewportCell.editText`); false by default. */
  editText?: boolean
}

/**
 * A cell of a window (`Session.getViewport`): its value as `getResult()` gives it, with the cell's whole look resolved, its
 * style and number format as rows and columns give them, what its conditional formats show, its note and link.
 */
export interface ViewportCell extends Partial<CellDisplay> {
  value: CellValue
  /** The canonical formula text, as `CellResult.formula`. */
  formula?: string
  /** With `formulaLocal`: the formula as Excel shows it in the book's locale (`Range.FormulaLocal` without `=`), inside its table as Excel shows it there (`[@Qty]`). */
  formulaLocal?: string
  spill?: string
  spillFill?: CellValue
  spillOf?: string
  /** On the top-left cell of an array formula of the old kind, its area (`CellInput.array`). */
  array?: string
  meta?: JsonValue
  /** The cell's own style, an index of the book's `styles` as `CellView.style`; its object is in `Viewport.styles`. */
  style?: number
  /**
   * The style the cell shows, an index of the book's `styles` (its object in `Viewport.styles`): its own, else its row's,
   * else its column's. It is always one of the table's: a cell where a styled row crosses a styled column shows the row's,
   * as Excel shows a file's empty cell there, and the two merged are a style of the cell's own (`Session.patchCells`).
   */
  effectiveStyle?: number
  /** The number format the cell shows by, as a code: its own, else its row's, else its column's, a built-in format by the code of the book's locale (14 is `dd.mm.yyyy` in ru-RU). */
  numberFormat?: string
  /** Where `numberFormat` comes from. */
  numberFormatSource?: 'cell' | 'row' | 'column'
  /** What the sheet's conditional formats give the cell by the current values; with a number format of theirs, `text` and `color` show the value by it. */
  appearance?: ConditionalAppearance
  note?: CellNote
  link?: CellLink
  /** With `diagnostics`, on a cell with an error or a formula a circular reference stops: as `Session.getCellIssue`. */
  issue?: CellIssueDetail
  /**
   * With `editText`: the text the cell opens with for editing, as Excel's formula bar shows it: a formula with `=` in the
   * book's locale (`formulaLocal`), in every cell of an old-kind array formula's area its formula (without the braces
   * the bar shows around it), a value as `entryText` writes it by the cell's number format, and empty for a cell
   * without a value or formula of its own (an empty one, one a dynamic array spilled into).
   */
  editText?: string
}

/** A column of a window: its number (A is 1), its width as `ColumnInput.width`, Excel's default where it has none. */
export interface ViewportColumn {
  column: number
  width: number
  hidden?: true
}

/** A row of a window: its number on the sheet, from 1, its height in points, Excel's default (15) where it has none. */
export interface ViewportRow {
  row: number
  height: number
  hidden?: true
}

/** A table that meets a window (`Viewport.tables`): its name, its range and which of its rows are a header and totals. */
export interface ViewportTable {
  name: string
  /** The whole table, its header and totals rows included (`TableInput.range`). */
  range: string
  /** Its first row is a header row (`TableInput.headers`, true by default). */
  headers: boolean
  /** Its last row is a totals row (`TableInput.totals`, false by default). */
  totals: boolean
}

/** What a grid needs to draw a window of a sheet (`Session.getViewport`). */
export interface Viewport {
  /**
   * The rectangle given (`B2:D9`): the one asked for, or as far as its first rows (of `ViewportOptions.rows`, with them)
   * that fit when it holds more than `maxCells` cells.
   */
  range: string
  /** The window asked for held more than `maxCells` cells: `range` is the part given. */
  truncated: boolean
  /** By address, row by row: every cell with something to show, an empty one with its row's or column's style or number format, an appearance of a conditional format, a note or a link too. */
  cells: Record<string, ViewportCell>
  /** The styles the window's cells have or show (`ViewportCell.style`, `effectiveStyle`), by their index in the book's `styles`: copies, the answer's own. */
  styles: Record<number, CellStyle>
  /** The columns of `range` left to right, or those of `ViewportOptions.columns`. */
  columns: ViewportColumn[]
  /** The rows of `range` top to bottom, or those of `ViewportOptions.rows`. */
  rows: ViewportRow[]
  /** The merged areas that meet the window's cells, whole. */
  merges: string[]
  /** The tables that meet the window's cells, whole, in the sheet's order. */
  tables: ViewportTable[]
  /** The sheet's frozen panes (`SheetInput.frozen`): its first `rows` rows and `columns` columns stay in view, 0 for none. */
  frozen: { rows: number; columns: number }
  /** Where the sheet's cells end, as `Session.getUsedRange`. */
  usedRange: SheetExtent
  /** The session's revision the window is of (`Session.revision`): a host takes the changes of later edits over it (`EditResult.revision`). */
  revision: number
}

/** `viewportOf`'s options: `getViewport`'s, and `appearance`. */
export interface SavedViewportOptions extends ViewportOptions {
  /**
   * Each cell also with the `appearance` its sheet's conditional formats give it, and its `text` and `color` by a number
   * format of theirs, as `getViewport` gives them in a session of the book on the result's clock; false by default. A
   * result holds no appearance: the rules are computed over the result's values, its formulas are not (TODAY in a rule
   * is the result's day). The first window with it over a book and a result reads all their values, the later ones
   * cost what the window and its rules read.
   */
  appearance?: boolean
}

/**
 * A cell of a window over a saved result (`viewportOf`): as `ViewportCell`, its value, `text` and `color` as the result
 * has them; without `text` and `color` when the result was computed without the display. The `appearance` of
 * conditional formats only with `SavedViewportOptions.appearance`: without it the cell's `text` shows the value by the
 * cell's own format, not a rule's.
 */
export type SavedViewportCell = ViewportCell

/** A window over a saved result (`viewportOf`): as `Viewport`, without a session's revision, its cells `SavedViewportCell`. */
export interface SavedViewport extends Omit<Viewport, 'revision' | 'cells'> {
  cells: Record<string, SavedViewportCell>
}

/**
 * A cell's data validation (`Session.getValidation`): the rule over it, the rule's formulas as at the cell (relative
 * references moved from the rule's first cell), and for a list the items its dropdown offers.
 */
export interface CellValidation {
  rule: DataValidation
  formula1?: string
  formula2?: string
  /**
   * A list's items in order, computed in the book as it is now: a literal list's items as written (`text`, and `value` the
   * same text), or the cells of the reference its formula gives, each with its value and the text it shows by its number
   * format (blank cells past the last value left out). An error where the formula gives no reference or list. `input` is
   * what Excel writes into the cell when the item is picked from its dropdown, for `setCells`, which keeps the cell's
   * format: a cell's value as it is (`001` stays text, a number shown `50%` is 0.5), a literal item read as typed in
   * en-US, the locale a file holds the list in (`1/2` is the 2nd of January, `d-mmm`), null for a blank cell. Excel
   * checks no rule on a picked item.
   */
  list?: { value: CellValue; text: string; input: CellInput | null }[] | { error: ErrorCode }
}

/** Where a sheet's cells end (`Session.getUsedRange`): everything lies in its first `rows` rows and `columns` columns; 0 and 0 for an empty sheet. */
export interface SheetExtent {
  rows: number
  columns: number
}

/** The way Control and an arrow key go from a cell (`Session.getEnd`). */
export type EndDirection = 'up' | 'down' | 'left' | 'right'
