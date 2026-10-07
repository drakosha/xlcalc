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
   * shown in the book's locale (`CellResult.text`). Without it a cell shows its row's format, else its column's.
   */
  format?: string
  /** A style of the book's `styles`, by index. Without it a cell shows its row's style, else its column's. */
  style?: number
  /** Arbitrary cell data (for example ids of source rows), passed through untouched. */
  meta?: JsonValue
}

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
  format?: string
}

/** Rows of a sheet, as xlsx `<row>`: the height, and the style and format of their cells that have none. */
export interface RowInput {
  /** Points, 0 to 409. */
  height?: number
  style?: number
  format?: string
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
}

export interface NameInput {
  /** Latin letters, digits, `_` and `.`; not a cell address, not R1C1, not TRUE/FALSE. */
  name: string
  /** A reference with a sheet: `'Data'!$D$2:$D$3401`, `'Data'!$D:$D`. */
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
}

/**
 * `CALC`: Excel's #CALC!, an empty array (`FILTER` that keeps nothing) or a text over 32,767 characters (`CONCAT`).
 * `SPILL`: an array that cannot spill from its formula's cell (its range is not empty, meets a merged area or
 * leaves the sheet).
 */
export type ErrorCode = 'DIV0' | 'NA' | 'NAME' | 'NUM' | 'REF' | 'VALUE' | 'CYCLE' | 'SPILL' | 'CALC' | 'ERROR'

/** A computed value. Dates come out as serial numbers, numbers are rounded to 15 significant digits. */
export type CellValue = Scalar | { error: ErrorCode }

/** Formula check codes. `unsupported`: an external function called with arguments that are not literals. */
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
  format?: string
  /** A tag of the package's locales (`'en-US'`, the default, or `'ru-RU'`). */
  locale?: string
  /** The column's width as `ColumnInput.width` (several columns' widths added for a merged area); the default width without it. */
  width?: number
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
}

/** A cell's value, and by default (`display`) what the cell shows: `text` and `color`, as `CellDisplay`. */
export interface CellResult extends Partial<CellDisplay> {
  /** The value: numbers stay numbers whatever the format shows. */
  value: CellValue
  /** Canonical formula text (sheet always quoted, function names upper-case). */
  formula?: string
  format?: string
  meta?: JsonValue
  /** A formula whose array spilled: the range it fills on its sheet (`B2:D9`, this cell at the top-left). */
  spill?: string
  /**
   * The value of the cells of `spill` that the result does not list: only for an array of more than 65,536
   * values whose cells past some row and column all hold one value (`A:A*2`: 0 below the data), which are
   * then left out of `cells` and of the hash's lines but for one line of the range.
   */
  spillFill?: CellValue
  /** A cell an array spilled into: the address of the array's formula on the same sheet (an empty element is 0). */
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
}

export type ValidateResult =
  | { ok: true; canonical: string; functions: string[]; sheets: string[]; names: string[]; externals: string[] }
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
 * default). With `display` (the default) results and changes carry what the cells show.
 */
export type SessionOptions = ({ clock?: Clock; live?: false } | { live: true }) & { externals?: Externals; execution?: Execution; display?: boolean }

/** A rectangle on a sheet: `B2`, `B2:D9`, whole columns `B:D` or whole rows `2:9` (no `$`, no sheet). */
export interface RangeRef {
  sheet: string
  range: string
}

/**
 * One structure operation (`restructure`). Rows are 1-based numbers and columns are letters, as in A1
 * addresses; `count` is 1 by default; sheets are matched case-insensitively.
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
   * (#SPILL!); clearing it lets the array spill again.
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
   * Column widths, row heights, row and column styles and formats, merged areas; applied atomically. Merged
   * areas stop arrays spilling over them: the formulas whose spill range an added or removed area meets are
   * recomputed and the changed values returned; nothing else is recomputed. With `display`, the cells whose shown
   * text or colour a width, a format or a merged area changed are returned too.
   */
  setLayout(edits: LayoutEdit[]): CellChange[]
  /**
   * The session's book with every change so far: what a host saves as a version. Its objects are the
   * session's own and later calls change them, so do not modify it; serialize or copy what must stay.
   */
  book(): BookInput
  explain(sheet: string, cell: string, opts?: ExplainOptions): Explanation
  result(): ComputeResult
  close(): void
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
