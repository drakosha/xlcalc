/**
 * The formula engine contract (subpath `./contract`): types only, no engine code. Any engine
 * implementation (currently formulas compiled to closures or to JS text) sits behind this contract; the
 * package's behaviour tests are its acceptance suite.
 *
 * A formula is stored in the `formula` field without the leading `=`, separately from the value:
 * a string from data never becomes a formula. Function names are English, arguments are separated
 * by commas, the decimal separator is a dot.
 */

export type Scalar = number | string | boolean | null

export type JsonValue = Scalar | JsonValue[] | { [key: string]: JsonValue }

/** An input cell. A date is passed explicitly (`{ date: '2025-01-31' }`) and becomes an Excel serial number. */
export interface CellInput {
  value?: Scalar | { date: string }
  /** Formula text without `=`. When present, `value` is ignored. */
  formula?: string
  /**
   * Number format, an Excel format code as an xlsx file stores it (en-US syntax: `#,##0.00`, `dd.mm.yyyy`, `[Red]`),
   * shown in the book's locale (`CellResult.text`), or a built-in format by its number (`NumberFormat`). Without it
   * a cell shows its row's format, else its column's.
   */
  format?: NumberFormat
  /** A style of the book's `styles`, by index. Without it a cell shows its row's style, else its column's. */
  style?: number
  /** Arbitrary cell data (for example ids of source rows), passed through untouched. */
  meta?: JsonValue
  /**
   * An array formula of the old kind (entered with Ctrl+Shift+Enter; in xlsx `<f t="array" ref="B2:C5">` without
   * dynamic array metadata), on the top-left cell of its area: the area (`B2:C5`, or the cell alone, `B2`). The
   * formula is computed as an array, with no implicit intersection, and fills the area, which it never leaves:
   * a smaller array is repeated when it has one row or one column, else the cells past it are #N/A, a larger one is
   * cut, one value fills every cell; nothing spills and nothing blocks it. `ROW()` and `COLUMN()` give the rows and
   * columns of the area. The other cells of the area hold no input, and a session refuses to change some of them
   * without the rest (`Session.setCells`). Only with `formula`.
   */
  array?: string
}

/**
 * A number format: a format code as an xlsx file stores it, or the number of one of Excel's built-in formats, 0 to
 * 49, as a file stores a cell's format with no code of its own (`numFmtId`). Excel shows some built-in formats by the
 * locale: 14 is the locale's short date (`m/d/yyyy` in en-US, `dd.mm.yyyy` in ru-RU), 5 to 8 and 37 to 44 its money
 * and number formats, 15 to 22 its other dates and times. A code is shown as written whatever the locale.
 */
export type NumberFormat = string | number

/** A colour, `#RRGGBB`. */
export type Color = string

/** A font; what a style leaves out is the book's default font. */
export interface FontStyle {
  name?: string
  /** Points, 1 to 409. */
  size?: number
  bold?: boolean
  italic?: boolean
  underline?: 'single' | 'double'
  strike?: boolean
  color?: Color
}

/** One side of a cell's border; the kinds are xlsx's. */
export interface BorderLine {
  style: 'thin' | 'medium' | 'thick' | 'dotted' | 'dashed' | 'double' | 'hair' | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot'
  color?: Color
}

export interface AlignmentStyle {
  horizontal?: 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed'
  vertical?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed'
  wrap?: boolean
  /** Indent level, 0 to 250. */
  indent?: number
}

/**
 * A cell style of the book's table (`BookInput.styles`), as an xlsx cell format without its number format
 * (that is the cell's `format`). Styles do not take part in computation or in the values hash.
 */
export interface CellStyle {
  font?: FontStyle
  /** A solid fill. */
  fill?: { color: Color }
  border?: { top?: BorderLine; right?: BorderLine; bottom?: BorderLine; left?: BorderLine }
  alignment?: AlignmentStyle
}

/** Columns of a sheet, as xlsx `<col>`: the width, and the style and format of their cells that have none. */
export interface ColumnInput {
  /**
   * In characters of the default font's digit with the cell's padding, as xlsx stores it, 0 to 255. Without it a
   * column has Excel's default width (8 digits and the padding; Excel for Windows shows it as 8.43).
   */
  width?: number
  style?: number
  format?: NumberFormat
  /** Hidden (xlsx `hidden`); a column of width 0 is hidden too. SUBTOTAL and AGGREGATE count the cells of hidden columns all the same, as Excel does. */
  hidden?: boolean
}

/** Rows of a sheet, as xlsx `<row>`: the height, and the style and format of their cells that have none. */
export interface RowInput {
  /** Points, 0 to 409. */
  height?: number
  style?: number
  format?: NumberFormat
  /**
   * Hidden (xlsx `hidden`), by hand or by a filter: SUBTOTAL with codes 101–111 and AGGREGATE with options 1, 3, 5
   * and 7 skip the row's cells, and on a `filtered` sheet SUBTOTAL with codes 1–11 too, as Excel does. A row of
   * `height` 0 is hidden too, as Excel takes it.
   */
  hidden?: boolean
}

export interface SheetInput {
  name: string
  /** Cells by A1 address (`B12`). */
  cells: Record<string, CellInput>
  /** Columns by letters, one (`C`) or a span (`B:D`) as xlsx `<col min max>`; spans do not overlap. */
  columns?: Record<string, ColumnInput>
  /** Rows by number, one (`5`) or a span (`5:9`); spans do not overlap. */
  rows?: Record<string, RowInput>
  /** Merged areas (`B2:D3`), as xlsx `<mergeCell>`: two cells or more each, none overlapping another. */
  merges?: string[]
  /**
   * A filter of the sheet has criteria, its autofilter's or a table's (Excel's filter mode): every hidden row of the
   * sheet counts as one a filter hid, which SUBTOTAL skips with codes 1–11 too and a copy leaves out, as Excel does
   * whatever hid the row.
   */
  filtered?: boolean
  /** The sheet's tables (Excel's list objects); they do not overlap, and no merged area or array area of several cells meets one. */
  tables?: TableInput[]
}

/**
 * A table (Excel's list object): a rectangle of a sheet with a header row of column names, data rows and a totals row.
 * Formulas refer to its parts by name with structured references (`Sales[Amount]`, `SUM(Sales[[#Totals],[Qty]])`),
 * which follow the table as rows and columns are inserted and deleted in it. A formula stores a reference to its own
 * row as an xlsx file does, `Sales[[#This Row],[Amount]]`; `Sales[@Amount]`, and `[@Amount]` inside the table, are read too.
 */
export interface TableInput {
  /**
   * Excel's rules for names: a letter of any script, `_` or `\` first, then letters, digits, `_` and `.`, at most 255
   * characters; not a cell address (`T2`) or R1C1 (`R`, `C`, `R1C1`), not TRUE or FALSE. Tables and defined names share
   * one namespace: no other table and no defined name of the book has the name, in any case.
   */
  name: string
  /** The table's rectangle, its header and totals rows included (`B2:E10`), with one data row at least. */
  range: string
  /** The first row is a header row (true by default): its cells hold the columns' names as text, and no formulas. */
  headers?: boolean
  /** The last row is a totals row (false by default). */
  totals?: boolean
  /** The columns from left to right, one for each column of the range. */
  columns: TableColumn[]
}

export interface TableColumn {
  /** 1 to 255 characters, unique in the table in any case; the header cell's text. */
  name: string
  /**
   * The formula of a calculated column (stored text, `Sales[[#This Row],[Qty]]*2`), which a row inserted into the table
   * gets in this column, as Excel fills it. The column's cells hold their own formulas.
   */
  formula?: string
  /**
   * What the column's cell of the totals row is, as an xlsx file records it: one of Excel's totals functions, a
   * SUBTOTAL over the column that the cell holds as its formula (`custom`: a formula of its own), or a label. The cell
   * computes what it holds.
   */
  totals?: { function: TotalsFunction } | { label: string }
}

/** Excel's totals functions of a table column (xlsx `totalsRowFunction`). */
export type TotalsFunction = 'sum' | 'average' | 'count' | 'countNums' | 'max' | 'min' | 'stdDev' | 'var' | 'custom'

export interface NameInput {
  /**
   * Excel's rules for names, a table's too: a letter of any script, `_` or `\` first, then letters, digits, `_` and `.`,
   * at most 255 characters; not a cell address (`Q1`) or R1C1 (`R`, `C`, `R2C3`), not TRUE or FALSE (`Ставка`,
   * `НДС_2025` and `АБ12` are names). Compared in any case: `ставка` is the name `Ставка`.
   */
  name: string
  /**
   * What the name refers to: a formula without `=`, as Excel's "Refers to" and xlsx `<definedName>` hold it. A
   * reference with a sheet (`'Data'!$D$2:$D$3401`, `'Data'!$D:$D`), a constant (`0.2`, `"text"`, `{1,2,3}`), a
   * formula (`'Data'!$A$1*2`, `SUM('Data'!$A:$A)`, `INDEX('Data'!$A$1:$C$9,2,0)`, which is a reference), a union of
   * references on one sheet (`'Data'!$A$1:$A$3,'Data'!$C$1:$C$3`). Every reference has its sheet. A reference
   * without `$` is relative to A1, as a file stores it: used in a cell it moves by that cell's offset from A1,
   * wrapping round the sheet's edges (`'Data'!B1` used in C5 is `'Data'!D5`, `'Data'!A1048576` is the row above).
   * Names in it are the global ones (a sheet's own as `'Sheet'!name`), as a file stores a name's formula. The formula
   * is computed where the name is used, as part of that formula.
   */
  ref: string
  /** The sheet where the name is visible; a local name shadows a global one. Without it the name is global. */
  scope?: string
}

export interface BookInput {
  /** Excel's rules for sheet names: 1 to 31 characters, none of `: \ / ? * [ ]`, no apostrophe at either end. */
  sheets: SheetInput[]
  names?: NameInput[]
  /** The cell styles that cells, rows and columns refer to by index. */
  styles?: CellStyle[]
  /**
   * The locale the book is computed in, a BCP 47 tag of one of the package's locales: `'en-US'` (the
   * default) or `'ru-RU'`. It changes computed values as Excel's regional settings do (text read as a
   * number or a date, numbers and booleans turned into text, `TEXT` format codes, currency), not the formula
   * text: formulas stay invariant, and `toLocal` shows them in a locale. An unknown tag is an `InputError`.
   */
  locale?: string
  /**
   * The date system, as xlsx `workbookPr date1904`: 1900 (the default) counts days from 1900-01-00 with Excel's
   * phantom 1900-02-29, 1904 from 1904-01-01, serial 0. It changes the serial numbers of the dates typed in and
   * computed, TODAY and NOW, date text and how dates show: a 1904 book shows a negative date or time with a minus
   * (`-12:00`). It goes into both hashes only when it is 1904.
   */
  dateSystem?: 1900 | 1904
  /**
   * Iterative calculation of circular references, as Excel's setting and xlsx `calcPr iterate`. Without it a cycle
   * gives `#CYCLE!`. With it the cells of every cycle and the cells that depend on them are computed from 0, again
   * and again, sheet by sheet in the order of the sheets' names (in any case), row by row. That stops after
   * `maxIterations` rounds (100 by default, 1 to 32,767) or when no value changes by `maxChange` or more (0.001 by
   * default). Each computation starts from 0 again, so a session gives what a full computation gives. Excel goes on
   * from the values it holds instead.
   */
  iteration?: { maxIterations?: number; maxChange?: number }
}

/**
 * `CALC`: Excel's #CALC!, an empty array (`FILTER` that keeps nothing) or a text over 32,767 characters (`CONCAT`).
 * `SPILL`: an array that cannot spill from its formula's cell (its range is not empty, meets a merged area or
 * leaves the sheet). `NULL`: Excel's #NULL!, ranges that do not meet under the intersection operator (`A1:A3 C1:C3`).
 * `ERROR`: a formula that does not compute (it does not parse, say).
 */
export type ErrorCode = 'DIV0' | 'NA' | 'NAME' | 'NUM' | 'REF' | 'VALUE' | 'NULL' | 'CYCLE' | 'SPILL' | 'CALC' | 'ERROR'

/** A computed value. Dates come out as serial numbers, numbers are rounded to 15 significant digits. */
export type CellValue = Scalar | { error: ErrorCode }

/**
 * Formula check codes. `unsupported`: an external function called with arguments that are not literals.
 * `unknown_name` covers tables too: an unknown table or column, a structured reference without a table outside one.
 */
export type CheckCode =
  | 'empty'
  | 'syntax'
  | 'too_long'
  | 'too_deep'
  | 'external_ref'
  | 'unknown_function'
  | 'unknown_sheet'
  | 'unknown_name'
  | 'unsupported'

/** The moment (ISO) and the random seed that TODAY, NOW, RAND and RANDBETWEEN are computed from. */
export interface Clock {
  now: string
  seed: number
}

export interface CellRef {
  sheet: string
  cell: string
}

/**
 * External functions (host functions evaluated outside the book): their values are resolved in
 * advance, in one request, and passed to the computation. The key of a call is its canonical text
 * (`BALANCE("2025-01")`); arguments are literals only.
 */
export interface Externals {
  functions: string[]
  values?: Record<string, CellValue>
}

/**
 * How formulas run; values, issues and hashes are the same either way. `closures` (the default) builds each
 * formula as a tree of JavaScript closures and generates no code, so it works under a Content Security
 * Policy without 'unsafe-eval'. `eval` generates JavaScript text and compiles it with `new Function`, which
 * such a policy blocks (computing then throws an `EvalError`).
 */
export type Execution = 'closures' | 'eval'

/** A value as a cell shows it: its text and the colour of its format's section. */
export interface CellDisplay {
  /**
   * What Excel shows in the cell (VBA's `Range.Text`): the value by its number format in the book's locale, General
   * narrowed to the column's width, `#` across the column for a number that does not fit, fills `*x` repeated
   * across it; text as it is (it overflows into the next cells). Derived from the value, the format, the locale and
   * the width: not part of the values hash.
   */
  text: string
  /** The colour of the format's section (`[Red]`, `[Color10]`), `#RRGGBB`; absent without one. */
  color?: Color
}

/** `displayValue`'s options: where the value is shown. */
export interface DisplayOptions {
  /** The number format; General without it. */
  format?: NumberFormat
  /** A tag of the package's locales (`'en-US'`, the default, or `'ru-RU'`). */
  locale?: string
  /** The column's width as `ColumnInput.width` (several columns' widths added for a merged area); the default width without it. */
  width?: number
  /** The date system a date shows by, as `BookInput.dateSystem`; 1900 without it. */
  dateSystem?: 1900 | 1904
}

/**
 * Size limits of a computation. A book over them is not computed: the call throws `CalcError` with code `limit`, so a
 * result is either exact or absent and the values hash never depends on the limits.
 */
export interface CalcLimits {
  /** Formulas in the book. */
  formulas?: number
  /** Values of one range read whole or one array that a formula makes, checked before they are made. */
  cells?: number
  /**
   * Values one spill may fill (its stored cells: the cells past a large array's last distinct value are not stored),
   * checked before they are written; `cells` when not given.
   */
  spill?: number
}

export interface ComputeOptions {
  /**
   * The clock for TODAY, NOW, RAND and RANDBETWEEN. By default the current moment and a random seed;
   * pass the `clock` of a result to reproduce that result exactly.
   */
  clock?: Clock
  externals?: Externals
  /** `closures` by default. */
  execution?: Execution
  /** Each cell's shown text and colour (`CellResult.text`, `color`); true by default, false computes values alone. */
  display?: boolean
  /** None by default. */
  limits?: CalcLimits
}

/** A cell's value, and by default (`display`) what the cell shows: `text` and `color`, as `CellDisplay`. */
export interface CellResult extends Partial<CellDisplay> {
  /** The value: numbers stay numbers whatever the format shows. */
  value: CellValue
  /** Canonical formula text (sheet always quoted, function names upper-case). */
  formula?: string
  format?: NumberFormat
  meta?: JsonValue
  /**
   * A formula whose array spilled: the range it fills on its sheet (`B2:D9`, this cell at the top-left); for an array
   * formula of the old kind (`CellInput.array`), its area.
   */
  spill?: string
  /**
   * The value of the cells of `spill` that the result does not list: only for an array of more than 65,536
   * values whose cells past some row and column all hold one value (`A:A*2`: 0 below the data), which are
   * then left out of `cells` and of the hash's lines but for one line of the range.
   */
  spillFill?: CellValue
  /**
   * A cell an array spilled into, or a cell of an old-kind array formula's area: the address of the array's formula
   * on the same sheet (an empty element is 0).
   */
  spillOf?: string
}

/** A cell with an error and its reason: an invalid formula (with the check code and position) or a computation error. */
export interface CellIssue extends CellRef {
  error: ErrorCode
  reason: string
  check?: CheckCode
  /** Position in the formula text for `check`. */
  at?: number
}

export interface NameIssue {
  name: string
  scope?: string
  reason: string
}

export interface ComputeResult {
  sheets: { name: string; cells: Record<string, CellResult> }[]
  issues: CellIssue[]
  nameIssues: NameIssue[]
  /** Cells that depend on volatile functions (directly or through other cells). */
  volatile: CellRef[]
  /** The clock the book was computed with (for a live session, the latest one). */
  clock: Clock
  /** sha256 of the canonical text of the values (spilled cells too) together with the engine configuration version. */
  hash: string
  /** sha256 of the input (sheets, cells, names, and styles, columns, rows and merged areas when there are any; order does not matter). */
  inputHash: string
  configVersion: string
  ms: number
}

export interface ValidateContext {
  /** Sheets of the book; without them sheets are not checked. */
  sheets?: string[]
  /** Names of the book; without them names are not checked. */
  names?: string[]
  externalFunctions?: string[]
  /**
   * Tables of the book with their columns; without them structured references are not checked. With them the canonical
   * text spells tables and columns as they are named, and a bare table name is a table, not a name.
   */
  tables?: { name: string; columns: string[] }[]
}

export type ValidateResult =
  /** `tables`: the tables the formula's structured references name (and bare table names, with `ctx.tables`). */
  | { ok: true; canonical: string; functions: string[]; sheets: string[]; names: string[]; externals: string[]; tables: string[] }
  | { ok: false; code: CheckCode; message: string; at: number }

export type NameCheck = { ok: true } | { ok: false; message: string }

/** A cell edit; `input: null` clears the cell. */
export interface CellEdit extends CellRef {
  input: CellInput | null
}

/**
 * A changed cell, with `spill`, `spillFill` and `spillOf`, and with `display` its `text` and `color`, as in
 * `CellResult`. A formula whose spill range or `spillFill` changed is listed even when its own value did not; the
 * cells it no longer lists are listed when their value changed. With `display`, a cell whose shown text or colour
 * changed is listed too (a new format, a column's width), with its value.
 */
/** A cell of `Session.cells`: as `CellResult`, and the cell's own `style` (an index of the book's `styles`). */
export interface CellView extends CellResult {
  style?: number
}

export interface CellChange extends CellRef, Partial<CellDisplay> {
  value: CellValue
  spill?: string
  spillFill?: CellValue
  spillOf?: string
}

export interface ExplainOptions {
  /** Depth over precedents, 3 by default. */
  depth?: number
  /** Node limit, 200 by default. */
  maxNodes?: number
  /** Expand ranges into their non-empty cells (by default a range is a single node). */
  expand?: boolean
}

export interface ExplainNode {
  /** `'Sheet'!B3` or `'Sheet'!D:D`. */
  ref: string
  value?: CellValue
  formula?: string
  /** A cell an array spilled into: its formula's cell (`'Sheet'!B2`), which is the node's child. */
  spillOf?: string
  /** For a range: the number of non-empty cells. */
  cells?: number
  children: ExplainNode[]
}

export interface Explanation {
  root: ExplainNode
  nodes: number
  truncated: boolean
  /** The tree as indented text (for logs or an LLM agent). */
  text: string
}

/**
 * Session options. A session keeps one clock, `clock` or by default the moment of opening and a random
 * seed, until `setClock`. A live session (`live: true`) takes a fresh clock and a new seed on every
 * recomputation instead. Edits are compiled in the `execution` the session was opened with (`closures` by
 * default). With `display` (the default) results and changes carry what the cells show. `limits` bound every call;
 * a call over them throws `CalcError` and leaves the session as it was.
 */
export type SessionOptions = ({ clock?: Clock; live?: false } | { live: true }) & { externals?: Externals; execution?: Execution; display?: boolean; limits?: CalcLimits }

/** A rectangle on a sheet: `B2`, `B2:D9`, whole columns `B:D` or whole rows `2:9` (no `$`, no sheet). */
export interface RangeRef {
  sheet: string
  range: string
}

/**
 * One structure operation (`restructure`). Rows are 1-based numbers and columns are letters, as in A1
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
   * The sheet moved before the sheet `before`, else after the last one. A reference across sheets covers the sheets
   * between its ends in the new order; an end that moves leaves it, the next sheet inward taking its place.
   */
  | { op: 'moveSheet'; sheet: string; before?: string }
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

/** Why Excel would refuse an operation. */
export type StructureRefusal =
  /** Cells would leave the sheet (past XFD or row 1,048,576), or a paste does not fit. */
  | 'off_sheet'
  /** A new sheet name that breaks Excel's rules or is taken (sheet names compare case-insensitively). */
  | 'sheet_name'
  /** The only sheet of a book cannot be deleted. */
  | 'last_sheet'
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

export type RestructureResult =
  | {
      ok: true
      /** The new book; the input book is not changed, unchanged sheets and cells are shared. */
      book: BookInput
      /** Where a cell of the old book is now; null if it was deleted or overwritten. */
      mapCell(ref: CellRef): CellRef | null
    }
  | { ok: false; code: StructureRefusal; message: string }

/** `Session.restructure`: the operation's result and the values it changed. */
export type SessionRestructureResult =
  | {
      ok: true
      /** The session's book after the operation (as `book()` returns it). */
      book: BookInput
      mapCell(ref: CellRef): CellRef | null
      /**
       * Cells whose value differs from the one their old cell (`mapCell`) had, by new address; a cell no old
       * cell came to counts as empty before. An interface that moves its shown values by `mapCell` and applies
       * these has the new values. With `display`, cells whose shown text or colour differs are listed too.
       */
      changes: CellChange[]
    }
  | { ok: false; code: StructureRefusal; message: string }

/** Fields to change on columns or rows: a value sets the field, null removes it, a field left out stays. */
export type LinePatch<T> = { [K in keyof T]?: T[K] | null }

/** A change of a sheet's layout. */
export type LayoutEdit =
  /** Columns `C` or `B:D`; the patch applies to each of them. */
  | { sheet: string; columns: string; set: LinePatch<ColumnInput> }
  /** Rows `5` or `5:9`; the patch applies to each of them. */
  | { sheet: string; rows: string; set: LinePatch<RowInput> }
  /** All the merged areas of the sheet. */
  | { sheet: string; merges: string[] }
  /** The sheet's filter mode (`SheetInput.filtered`). */
  | { sheet: string; filtered: boolean }

/**
 * A session over a book: edits recompute only the cells that depend on them. The current clock of the
 * session (`clock`) is what a saved result should record: a full computation with it gives the same hash.
 */
export interface Session {
  readonly clock: Clock
  /**
   * Applied atomically with one recomputation; returns the changed values. An edit that keeps the cell's
   * value and formula (a new `style`, `format` or `meta`) recomputes nothing; with `display` a new format returns
   * the cell when its shown text changed. A value or formula put into a spill range stops the array spilling
   * (#SPILL!); clearing it lets the array spill again. Edits that change some cells of an old-kind array formula's
   * area (`CellInput.array`) but not all of them throw `InputError` and change nothing, as Excel refuses to change
   * part of an array; an edit of its top-left cell that keeps an array formula over the same area changes all of it.
   * An edit of a table's header cell renames its column, as Excel does: the structured references to it follow, a number
   * or a boolean becomes its text, a cleared cell names the column `Column1` (the first such name free; `Столбец1` in a
   * ru-RU book), and a name another column has already gives the column further right a number (`Price2`). A formula
   * in a header cell throws `InputError`.
   */
  setCells(edits: CellEdit[]): CellChange[]
  setNames(names: NameInput[]): CellChange[]
  /** Non-live sessions only. */
  setClock(clock: Clock): CellChange[]
  /** A structure operation on the session's book, recomputed; refused as `restructure` refuses it. */
  restructure(op: StructureOp): SessionRestructureResult
  /** Replaces the book's style table; every style that cells, rows and columns use must stay in it. Recomputes nothing; the shown text does not depend on styles. */
  setStyles(styles: CellStyle[]): void
  /**
   * Column widths, row heights, row and column styles and formats, hidden rows and columns, merged areas, the filter
   * mode; applied atomically. Merged areas stop arrays spilling over them: the formulas whose spill range an added
   * or removed area meets are recomputed. Hidden rows and the filter mode change what SUBTOTAL and AGGREGATE count:
   * their formulas over the rows concerned are recomputed. Nothing else is recomputed; the changed values are
   * returned. With `display`, the cells whose shown text or colour a width, a format or a merged area changed are
   * returned too.
   */
  setLayout(edits: LayoutEdit[]): CellChange[]
  /**
   * The session's book with every change so far: what a host saves as a version. Its objects are the
   * session's own and later calls change them, so do not modify it; serialize or copy what must stay.
   */
  book(): BookInput
  explain(sheet: string, cell: string, opts?: ExplainOptions): Explanation
  /**
   * The cells of a rectangle of a sheet (`B2:D9`, `B2`, whole columns `B:D`, whole rows `2:9`) as `result()` gives
   * them, by address: the same cells (spilled ones too, empty ones left out) with the same fields, and each cell's own
   * `style`. For a window over a sheet: it costs what the rectangle holds, not the book.
   */
  cells(sheet: string, range: string): Record<string, CellView>
  result(): ComputeResult
  /** The values hash `result()` gives, without building the result. */
  hash(): string
  close(): void
}

// ---------- asynchronous and threaded computation (`docs/async.md`) ----------

/** What a computation needs of an AbortSignal: the platform's `AbortSignal` is one. */
export interface AbortSignalLike {
  readonly aborted: boolean
  readonly reason?: unknown
  addEventListener(type: 'abort', listener: () => void, options?: { once?: boolean }): void
  removeEventListener(type: 'abort', listener: () => void): void
}

/**
 * Why an asynchronous call failed (`CalcError.code`). `soft_timeout`: its `timeoutMs` passed, seen by a check inside
 * the computation. `aborted`: its `signal` fired. `hard_timeout`: it ran past the pool's `hardTimeoutMs` and its
 * thread was terminated. `out_of_memory`: its thread reached `maxHeapMb`. `thread_failed`: its thread exited or threw
 * outside the computation. `restore_failed`: the session could not be restored on a thread and is closed. `limit`:
 * the book is over `CalcLimits`. `queue_full`: `maxQueue` tasks already wait. `too_many_sessions`: the pool holds
 * `maxSessions` sessions. `closed`: the calculator or the session is closed.
 */
export type CalcErrorCode =
  | 'soft_timeout'
  | 'aborted'
  | 'hard_timeout'
  | 'out_of_memory'
  | 'thread_failed'
  | 'restore_failed'
  | 'limit'
  | 'queue_full'
  | 'too_many_sessions'
  | 'closed'

/** Options of one asynchronous call. */
export interface CalcOptions {
  /** Stops the call: a waiting call at once, a running one at its next check. */
  signal?: AbortSignalLike
  /** A soft deadline: milliseconds of computing, counted from when the call starts computing (not while it waits). */
  timeoutMs?: number
}

export type AsyncComputeOptions = ComputeOptions & CalcOptions
/** `signal` and `timeoutMs` bound the opening (the first computation of the book). */
export type AsyncSessionOptions = SessionOptions & CalcOptions

/** `AsyncSession.restructure`: as `SessionRestructureResult`, without the book (`AsyncSession.book()` gives it). */
export type AsyncRestructureResult =
  | { ok: true; mapCell(ref: CellRef): CellRef | null; changes: CellChange[] }
  | { ok: false; code: StructureRefusal; message: string }

/**
 * A session computed in portions or in a thread: the methods of `Session`, as promises. Calls are taken in order,
 * each after the one before has finished. A call that fails (a deadline, an abort, a limit) changes nothing.
 */
export interface AsyncSession {
  /** The session's clock as of its last finished call. */
  readonly clock: Clock
  setCells(edits: CellEdit[], opts?: CalcOptions): Promise<CellChange[]>
  setNames(names: NameInput[], opts?: CalcOptions): Promise<CellChange[]>
  setClock(clock: Clock, opts?: CalcOptions): Promise<CellChange[]>
  restructure(op: StructureOp, opts?: CalcOptions): Promise<AsyncRestructureResult>
  setStyles(styles: CellStyle[]): Promise<void>
  setLayout(edits: LayoutEdit[], opts?: CalcOptions): Promise<CellChange[]>
  /** The session's book; from a thread, a copy. */
  book(): Promise<BookInput>
  explain(sheet: string, cell: string, opts?: ExplainOptions): Promise<Explanation>
  /** As `Session.cells`; from a thread only the rectangle comes back. */
  cells(sheet: string, range: string): Promise<Record<string, CellView>>
  result(opts?: CalcOptions): Promise<ComputeResult>
  /** As `Session.hash`: from a thread only the hash comes back. */
  hash(opts?: CalcOptions): Promise<string>
  close(): Promise<void>
}

/** A calculator: `createInlineCalc` (the same thread, in portions) or `createCalcPool` (threads). */
export interface Calc {
  compute(book: BookInput, opts?: AsyncComputeOptions): Promise<ComputeResult>
  openSession(book: BookInput, opts?: AsyncSessionOptions): Promise<AsyncSession>
  /** Refuses further calls (`closed`), the waiting and running ones too, and closes the sessions. */
  close(): Promise<void>
}

export interface InlineCalcOptions {
  /** Milliseconds of computing between two yields to the event loop; 10 by default. */
  sliceMs?: number
  /** The soft deadline of every call that gives none. */
  timeoutMs?: number
  /** The limits of every computation; a computation's own `limits` take their place field by field, the others stay. */
  limits?: CalcLimits
}

/** What the pool needs of a Web Worker (`CalcPoolOptions.spawn`). */
export interface WorkerLike {
  postMessage(message: unknown): void
  terminate(): void
  addEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { data?: unknown; message?: string }) => void): void
}

export interface CalcPoolOptions {
  /** Threads; 2 by default. */
  threads?: number
  /** Node: the heap of each thread (`resourceLimits.maxOldGenerationSizeMb`); not available in a browser. */
  maxHeapMb?: number
  /** A task that runs longer, from when it reaches its thread, terminates the thread (`hard_timeout`). */
  hardTimeoutMs?: number
  /** The soft deadline of every call that gives none. */
  timeoutMs?: number
  /** Tasks waiting for a thread, all threads together; one more is refused (`queue_full`). */
  maxQueue?: number
  /** Open sessions of the pool; one more is refused (`too_many_sessions`). */
  maxSessions?: number
  /** Sessions a thread holds; past it the session called least recently leaves the thread and comes back on its next call. */
  maxSessionsPerThread?: number
  /**
   * The limits of every computation; a computation's own `limits` take their place field by field, the others stay. With `maxHeapMb`, `cells` is `maxHeapMb · 2^20 / 64` where not
   * given, and `spill` the smaller of `cells` and `maxHeapMb · 2^20 / 1024` where not given: one allocation larger than
   * what a thread's heap has left (a large array, the tables of a large spill's cells) stops the whole process, not the
   * thread.
   */
  limits?: CalcLimits
  /** Calls kept in a session's journal before the pool takes a new snapshot of its book; 64 by default. */
  snapshotEvery?: number
  /** Milliseconds of computing between two yields inside a thread, where a cancel is read; 50 by default. */
  sliceMs?: number
  /** Starts a thread in a browser, for a bundler that serves the thread's script elsewhere. */
  spawn?: () => WorkerLike
}

export interface FormulaEngine {
  compute(book: BookInput, opts?: ComputeOptions): ComputeResult
  validateFormula(text: string, ctx?: ValidateContext): ValidateResult
  validateName(name: string): NameCheck
  openSession(book: BookInput, opts?: SessionOptions): Session
  restructure(book: BookInput, op: StructureOp): RestructureResult
  /** A value as a cell shows it, without a book: by a format, in a locale and a column's width. */
  displayValue(value: CellValue, opts?: DisplayOptions): CellDisplay
}

export const MAX_FORMULA_LENGTH = 8_192
export const MAX_FORMULA_DEPTH = 64
