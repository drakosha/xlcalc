/** A computation and what it reports: its options and limits, a cell's result and how it shows, the issues of cells and names, the checks of formulas and names, why a cell holds its value, a book with its result. */
import type { Locale } from '../locale/types.js'
import type { CellRef, CellValue, CheckCode, Clock, ErrorCode, JsonValue, RangeRef, Refusal } from './values.js'
import type { Book, ExternalBook, NumberFormat } from './book.js'
import type { Color } from './looks.js'

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

/**
 * Size limits of a computation and of an edit. A book over them is not computed and an edit over them is not made: the
 * call throws `CalcError` with code `limit`, so a result is either exact or absent and the values hash never depends on
 * the limits.
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
  /**
   * Cells one edit of a session writes with a value or a formula (a look alone goes to the sheet's looks), counted
   * before anything is written: a copy's source cells that hold one times the tiles of its target, a move's source cells
   * that hold one, the cells of `fillCellText`'s range and of `pasteText`'s rectangle, the cells `replaceInCells` enters
   * again, the edits of `setCells`, `replaceCells`, `setCellTexts` and `applyChanges`; their previews too. Operations
   * that write no more than the sheet holds (rows and columns inserted and deleted, `clearCells`) are not bounded by it.
   */
  writes?: number
}

/**
 * The values a file stores for its book's formulas, in the shape of a result: `readXlsx`'s `cached`, or a result of the
 * book. They are read by sheet name and address, a cell's `value` alone.
 */
export interface CachedValues {
  sheets: { name: string; cells: Record<string, { value: CellValue }> }[]
}

/** Options of `compute` and of a session: the clock, the way formulas run, limits and the values a file stored. */
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
  /**
   * The values the book's file stores (`readXlsx`'s `cached`), as Excel opens a file with a circular reference: a formula
   * the cycle stops (`CellInput.held`) without a `held` of its own holds the value stored for it, empty text where the
   * file stores none. Nothing else is read from them.
   */
  cached?: CachedValues
}

/** A cell's value, and by default (`display`) what the cell shows: `text` and `color`, as `CellDisplay`. */
export interface CellResult extends Partial<CellDisplay> {
  /** The value: numbers stay numbers whatever the format shows. */
  value: CellValue
  /**
   * Canonical formula text (sheet always quoted, function names upper-case), with sheets, tables and defined names
   * spelled as the book names them, as Excel shows them (`ставка*2` with the name `Ставка` is `Ставка*2`).
   */
  formula?: string
  /** The cell's own number format, as its input has it. */
  numberFormat?: NumberFormat
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

/**
 * A cell with an error and its reason: an invalid formula (with the check code and position) or a computation error; or
 * a formula a circular reference stops (`error: 'CYCLE'`), whose value is the one it holds (`CellInput.held`).
 */
export interface CellIssue extends CellRef {
  error: ErrorCode
  reason: string
  check?: CheckCode
  /** Position in the formula text for `check`. */
  at?: number
}

/** Why a defined name gives no value: the name, its scope (none for a global name) and the reason. */
export interface NameIssue {
  name: string
  scope?: string
  reason: string
  /**
   * The book has no such name: one Excel would not take, a duplicate, an unknown scope sheet, a formula that does not
   * parse or a reference without its sheet; a formula using the word finds no name. False: the name is in the book and
   * gives an error where it is used (an unknown sheet, table or function in its formula).
   */
  dropped: boolean
}

/**
 * What a result reports about its book, by `kind`: a cell with an error and its reason (`cell`, a `CellIssue`), or a
 * name the book left out or keeps with an error (`name`, a `NameIssue`).
 */
export type Diagnostic = ({ kind: 'cell' } & CellIssue) | ({ kind: 'name' } & NameIssue)

/** What `compute` gives: each sheet's cells with their values and shown text, the issues, the values' `hash`, `inputHash` and the clock it computed with. */
export interface ComputeResult {
  sheets: { name: string; cells: Record<string, CellResult> }[]
  /**
   * Every cell with an error and every formula a circular reference stops (`CellIssue`), sheet by sheet in the book's
   * order and row by row, then every name with a problem, in the order of the book's names.
   */
  diagnostics: Diagnostic[]
  /** Cells that depend on volatile functions (directly or through other cells). */
  volatile: CellRef[]
  /** The clock the book was computed with (for a live session, the latest one). */
  clock: Clock
  /** sha256 of the canonical text of the values (spilled cells too) together with the engine configuration version. */
  hash: string
  /** sha256 of the input (sheets, cells, names, and styles, columns, rows, merged areas, frozen panes, conditional formats, data validation and outline settings when there are any; order does not matter). */
  inputHash: string
  configVersion: string
  ms: number
}

/** What `validateFormula` checks a formula against: the book's sheets, names and tables; without them they are not checked. */
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
  /**
   * The table of the formula's cell: a structured reference without a table name is to it (`[@Qty]`), as in a cell of
   * the table, and the canonical text names it (`Sales[[#This Row],[Qty]]`). Without it such a reference is refused
   * with `tables` (`unknown_name`) and stays without a name otherwise.
   */
  table?: string
  /**
   * The book's links to other workbooks (`BookInput.externalBooks`): a reference to one is checked against its sheets
   * and names. A reference to a book the links do not hold is refused (`external_ref`), with or without them.
   */
  externalBooks?: ExternalBook[]
}

/** What `validateFormula` answers: the formula's canonical text and what it names, or why it does not read and where. */
export type ValidateResult =
  /** `tables`: the tables the formula's structured references name (and bare table names, with `ctx.tables`). */
  | { ok: true; canonical: string; functions: string[]; sheets: string[]; names: string[]; externals: string[]; tables: string[] }
  | (Refusal<CheckCode> & { at: number })

/**
 * Why Excel would not take a name (`validateName`). `chars`: it does not start with a letter of any script, `_` or `\`,
 * holds a character other than letters, digits, `_`, `.`, `\` and `?`, or is empty or longer than 255 characters.
 * `reference`: it reads as a cell address or an R1C1 reference (`Q1`, `R`, `R2C3`). `boolean`: TRUE or FALSE, or the
 * book's locale's word for one, in any case (`ИСТИНА` in ru-RU).
 */
export type NameRefusal = 'chars' | 'reference' | 'boolean'

/** What `validateName` answers: the name is good, or why not. */
export type NameCheck = { ok: true } | Refusal<NameRefusal>

/** The book whose rules `validateName` checks a name by. */
export interface NameOptions {
  /** The book's locale (`BookInput.locale`): a loaded locale's tag, or a locale object; en-US without one. */
  locale?: string | Locale
}

/** The parts of a book `validateNames` checks names against. */
export interface NamesContext {
  /** Sheets of the book: a name's scope must be one of them, and its formula's references to others give `#REF!`. */
  sheets: string[]
  /** Tables of the book with their columns, for the structured references of names' formulas. */
  tables?: { name: string; columns: string[] }[]
  /** The book's links to other workbooks, for the names' references to them (`ValidateContext.externalBooks`). */
  externalBooks?: ExternalBook[]
  externalFunctions?: string[]
  /** The book's locale (`BookInput.locale`): a loaded locale's tag, or a locale object; en-US without one. */
  locale?: string | Locale
}

/**
 * Why a cell holds an error, or a circular reference stops it (`Session.getCellIssue`): its code, reason, check code and
 * position as the result's `diagnostics` gives them, and `origin`, the cell where the error began, following the cells
 * its reason names (`from 'S'!A2`) to one whose error is its own; in a cycle, the first cell of the cycle the way
 * reaches, and so for a formula a cycle stops, through the cells it reads that the cycle stops too.
 */
export interface CellIssueDetail {
  error: ErrorCode
  reason: string
  check?: CheckCode
  at?: number
  origin: CellRef
}

/** Options of `explain`: how deep the tree goes and how many nodes it holds. */
export interface ExplainOptions {
  /** Depth over precedents, 3 by default. */
  depth?: number
  /** Node limit, 200 by default. */
  maxNodes?: number
  /** Expand ranges into their non-empty cells (by default a range is a single node). */
  expand?: boolean
  /** How many cells past `depth` the way from an error to the cell it began in goes on, 64 by default. */
  errorDepth?: number
}

/** A node of an explanation (`Explanation.root`): a cell or a range, its formula and value, and the nodes it reads. */
export interface ExplainNode {
  /** The local node's sheet and A1 cell or range for navigation; absent for another workbook. */
  location?: RangeRef
  /** `'Sheet'!B3` or `'Sheet'!D:D`. */
  ref: string
  value?: CellValue
  formula?: string
  /** A cell an array spilled into: its formula's cell (`'Sheet'!B2`), which is the node's child. */
  spillOf?: string
  /** For a range: the number of non-empty cells. */
  cells?: number
  /**
   * A cell with an error, or a formula a circular reference stops: the cell the error began in (`'Sheet'!A2`, the node's
   * own `ref` when it began there), as `CellIssueDetail.origin`. The way there is in the tree past `depth` too, and
   * through a range by the cell it came from.
   */
  origin?: string
  children: ExplainNode[]
}

/** What `explain` answers: why a cell holds its value, as a tree of what it reads and as text. */
export interface Explanation {
  root: ExplainNode
  nodes: number
  truncated: boolean
  /**
   * The tree as indented text (for logs, an LLM agent or a person) in the book's locale: in ru-RU its words, numbers,
   * formulas, booleans and errors are Russian (`План!B2 = 480  ← ОКРУГЛ(B1*1,2;0)`), and a sheet is in quotes only where
   * Excel puts them; the nodes keep the stored text. An error is written as the book's cells show it (`#DIV/0!`), with the
   * cell it began in (`(error from S!A2)`).
   */
  text: string
}

/**
 * A book and its result: a session's taken at one moment (`Session.getSnapshot`), or a book and what `compute` gave for
 * it, as `writeXlsx` takes them and `viewportOf` reads a version a host saved.
 */
export interface BookSnapshot {
  book: Book
  result: ComputeResult
}
