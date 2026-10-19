/** A computation and what it reports: its options and limits, a cell's result and how it shows, the issues of cells and names, the checks of formulas and names, why a cell holds its value, a book with its result. */
import type { Locale } from '../locale/types.js'
import type { CellRef, CellValue, CheckCode, Clock, ErrorCode, JsonValue, RangeRef, Refusal } from './values.js'
import type { Book, ExternalBook, NumberFormat } from './book.js'
import type { Color } from './looks.js'

/**
 * External functions (host functions evaluated outside the workbook): their values are resolved in
 * advance, in a single request, and passed to the calculation. The key of a call is its canonical text
 * (`BALANCE("2025-01")`); arguments must be literals.
 */
export interface Externals {
  functions: string[]
  values?: Record<string, CellValue>
}

/**
 * How formulas are executed; values, issues and hashes are the same either way. `closures` (the default) builds each
 * formula as a tree of JavaScript closures and generates no code, so it works under a Content Security
 * Policy without 'unsafe-eval'. `eval` generates JavaScript source and compiles it with `new Function`, which
 * such a policy blocks (calculation then throws an `EvalError`).
 */
export type Execution = 'closures' | 'eval'

/** A value as a cell displays it: its text and the color of its format's section. */
export interface CellDisplay {
  /**
   * What Excel displays in the cell (VBA's `Range.Text`): the value in its number format and the workbook's locale,
   * with General shortened to fit the column's width, `#` across the column for a number that does not fit, and fills
   * (`*x`) repeated across it; text as is (it overflows into the adjacent cells). Derived from the value, the format, the
   * locale and the width, so it is not part of the values hash.
   */
  text: string
  /** The color of the format's section (`[Red]`, `[Color10]`), `#RRGGBB`; absent when the section has none. */
  color?: Color
}

/**
 * Size limits for a calculation and for an edit. A workbook that exceeds them is not calculated, and an edit that
 * exceeds them is not made: the call throws `CalcError` with code `limit`, so a result is either exact or absent, and
 * the values hash never depends on the limits.
 */
export interface CalcLimits {
  /** Formulas in the workbook. */
  formulas?: number
  /** Values in a single range read as a whole, or in a single array a formula creates, checked before they are created. */
  cells?: number
  /**
   * Values a single spill may fill (its stored cells: the cells after a large array's last distinct value are not
   * stored), checked before they are written; `cells` when not given.
   */
  spill?: number
  /**
   * Cells a single session edit writes a value or a formula to (formatting alone is stored in the sheet's `looks`),
   * counted before anything is written: a copy's source cells that hold one, multiplied by the number of tiles in its
   * target; a move's source cells that hold one; the cells of `fillCellText`'s range and of `pasteText`'s rectangle; the
   * cells `replaceInCells` re-enters; the edits of `setCells`, `replaceCells`, `setCellTexts` and `applyChanges`; and
   * their previews. Operations that write no more than the sheet already holds (inserting and deleting rows and columns,
   * `clearCells`) are not limited by it.
   */
  writes?: number
}

/**
 * The values a file stores for its workbook's formulas, in the shape of a result: `readXlsx`'s `cached`, or a result
 * for the workbook. They are looked up by sheet name and address, and only a cell's `value` is read.
 */
export interface CachedValues {
  sheets: { name: string; cells: Record<string, { value: CellValue }> }[]
}

/** Options of `compute` and of a session: the clock, how formulas are executed, limits and the values a file stored. */
export interface ComputeOptions {
  /**
   * The clock for TODAY, NOW, RAND and RANDBETWEEN. By default, the current time and a random seed;
   * pass the `clock` of a result to reproduce that result exactly.
   */
  clock?: Clock
  externals?: Externals
  /** `closures` by default. */
  execution?: Execution
  /** Each cell's displayed text and color (`CellResult.text`, `color`); true by default, false calculates values only. */
  display?: boolean
  /** None by default. */
  limits?: CalcLimits
  /**
   * The values stored in the workbook's file (`readXlsx`'s `cached`), used the way Excel opens a file with a circular
   * reference: a formula stopped by the cycle (`CellInput.held`) that has no `held` of its own keeps the value stored for
   * it, or empty text when the file stores none. Nothing else is read from them.
   */
  cached?: CachedValues
}

/** A cell's value and, by default (`display`), what the cell displays: `text` and `color`, as in `CellDisplay`. */
export interface CellResult extends Partial<CellDisplay> {
  /** The value: numbers remain numbers whatever the format displays. */
  value: CellValue
  /**
   * Canonical formula text (sheet names always quoted, function names in uppercase), with sheets, tables and defined
   * names spelled as the workbook defines them, as Excel displays them (`ставка*2` with the name `Ставка` becomes
   * `Ставка*2`).
   */
  formula?: string
  /** The cell's own number format, as given in its input. */
  numberFormat?: NumberFormat
  meta?: JsonValue
  /**
   * A formula whose array spilled: the range it fills on its sheet (`B2:D9`, with this cell at the top left); for a
   * legacy array formula (`CellInput.array`), its range.
   */
  spill?: string
  /**
   * The value of the cells of `spill` that the result does not list: only for an array of more than 65,536
   * values whose cells beyond some row and column all hold the same value (`A:A*2`: 0 below the data). Those
   * cells are then left out of `cells`, and out of the hash's lines except for a single line for the range.
   */
  spillFill?: CellValue
  /**
   * A cell that an array spilled into, or a cell in a legacy array formula's range: the address of the array's formula
   * on the same sheet (an empty element is 0).
   */
  spillOf?: string
}

/**
 * A cell with an error and its reason: an invalid formula (with the check code and position) or a calculation error; or
 * a formula stopped by a circular reference (`error: 'CYCLE'`), whose value is the one it keeps (`CellInput.held`).
 */
export interface CellIssue extends CellRef {
  error: ErrorCode
  reason: string
  check?: CheckCode
  /** Position in the formula text for `check`. */
  at?: number
}

/** Why a defined name returns no value: the name, its scope (none for a global name) and the reason. */
export interface NameIssue {
  name: string
  scope?: string
  reason: string
  /**
   * The workbook has no such name: one Excel would reject, a duplicate, an unknown scope sheet, a formula that does not
   * parse, or a reference without its sheet; a formula that uses the word finds no name. False: the name is in the
   * workbook and returns an error where it is used (an unknown sheet, table or function in its formula).
   */
  dropped: boolean
}

/**
 * What a result reports about its workbook, by `kind`: a cell with an error and its reason (`cell`, a `CellIssue`), or a
 * name that the workbook dropped or keeps with an error (`name`, a `NameIssue`).
 */
export type Diagnostic = ({ kind: 'cell' } & CellIssue) | ({ kind: 'name' } & NameIssue)

/** What `compute` returns: each sheet's cells with their values and displayed text, the issues, the values' `hash`, `inputHash` and the clock it calculated with. */
export interface ComputeResult {
  sheets: { name: string; cells: Record<string, CellResult> }[]
  /**
   * Every cell with an error and every formula stopped by a circular reference (`CellIssue`), sheet by sheet in the
   * workbook's order and row by row, then every name with a problem, in the order of the workbook's names.
   */
  diagnostics: Diagnostic[]
  /** Cells that depend on volatile functions (directly or through other cells). */
  volatile: CellRef[]
  /** The clock the workbook was calculated with (for a live session, the latest one). */
  clock: Clock
  /** sha256 of the canonical text of the values (including spilled cells) together with the engine configuration version. */
  hash: string
  /** sha256 of the input (sheets, cells, names, and styles, columns, rows, merged areas, frozen panes, conditional formats, data validation and outline settings when there are any; order does not matter). */
  inputHash: string
  configVersion: string
  ms: number
}

/** What `validateFormula` checks a formula against: the workbook's sheets, names and tables; whatever is omitted is not checked. */
export interface ValidateContext {
  /** Sheets of the workbook; without them, sheets are not checked. */
  sheets?: string[]
  /** Names of the workbook; without them, names are not checked. */
  names?: string[]
  externalFunctions?: string[]
  /**
   * Tables of the workbook with their columns; without them, structured references are not checked. With them, the
   * canonical text spells tables and columns as they are named, and a bare table name refers to a table, not a defined
   * name.
   */
  tables?: { name: string; columns: string[] }[]
  /**
   * The table containing the formula's cell: a structured reference without a table name refers to it (`[@Qty]`), as in
   * a cell of the table, and the canonical text names the table (`Sales[[#This Row],[Qty]]`). Without it, such a
   * reference is rejected when `tables` is given (`unknown_name`), and otherwise stays without a table name.
   */
  table?: string
  /**
   * The workbook's links to other workbooks (`BookInput.externalBooks`): a reference to one of them is checked against
   * its sheets and names. A reference to a workbook not among the links is rejected (`external_ref`), with or without
   * them.
   */
  externalBooks?: ExternalBook[]
}

/** What `validateFormula` returns: the formula's canonical text and what it refers to, or why and where it fails to parse. */
export type ValidateResult =
  /** `tables`: the tables named by the formula's structured references (and bare table names, with `ctx.tables`). */
  | { ok: true; canonical: string; functions: string[]; sheets: string[]; names: string[]; externals: string[]; tables: string[] }
  | (Refusal<CheckCode> & { at: number })

/**
 * Why Excel would reject a name (`validateName`). `chars`: it does not start with a letter of any script, `_` or `\`,
 * contains a character other than letters, digits, `_`, `.`, `\` and `?`, or is empty or longer than 255 characters.
 * `reference`: it is parsed as a cell address or an R1C1 reference (`Q1`, `R`, `R2C3`). `boolean`: TRUE or FALSE, or
 * the workbook locale's word for either, in any case (`ИСТИНА` in ru-RU).
 */
export type NameRefusal = 'chars' | 'reference' | 'boolean'

/** What `validateName` returns: the name is valid, or why it is not. */
export type NameCheck = { ok: true } | Refusal<NameRefusal>

/** The workbook whose rules `validateName` checks a name against. */
export interface NameOptions {
  /** The workbook's locale (`BookInput.locale`): a loaded locale's tag, or a locale object; en-US without one. */
  locale?: string | Locale
}

/** The parts of a workbook that `validateNames` checks names against. */
export interface NamesContext {
  /** Sheets of the workbook: a name's scope must be one of them, and references in its formula to other sheets return `#REF!`. */
  sheets: string[]
  /** Tables of the workbook with their columns, for structured references in the names' formulas. */
  tables?: { name: string; columns: string[] }[]
  /** The workbook's links to other workbooks, for references to them in the names (`ValidateContext.externalBooks`). */
  externalBooks?: ExternalBook[]
  externalFunctions?: string[]
  /** The workbook's locale (`BookInput.locale`): a loaded locale's tag, or a locale object; en-US without one. */
  locale?: string | Locale
}

/**
 * Why a cell holds an error, or why a circular reference stops it (`Session.getCellIssue`): its code, reason, check code
 * and position as the result's `diagnostics` reports them, and `origin`, the cell where the error originated, found by
 * following the cells its reason names (`from 'S'!A2`) to one whose error is its own; in a cycle, the first cell of the
 * cycle that the path reaches, and likewise for a formula stopped by a cycle, through the cells it reads that the cycle
 * also stops.
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
  /** Maximum depth of precedents, 3 by default. */
  depth?: number
  /** Node limit, 200 by default. */
  maxNodes?: number
  /** Expand ranges into their non-empty cells (by default a range is a single node). */
  expand?: boolean
  /** How many cells beyond `depth` the path from an error to the cell where it originated may continue, 64 by default. */
  errorDepth?: number
}

/** A node of an explanation (`Explanation.root`): a cell or a range, its formula and value, and the nodes it reads. */
export interface ExplainNode {
  /** The sheet and A1 cell or range of a node in this workbook, for navigation; absent for an external workbook. */
  location?: RangeRef
  /** `'Sheet'!B3` or `'Sheet'!D:D`. */
  ref: string
  value?: CellValue
  formula?: string
  /** A cell that an array spilled into: its formula's cell (`'Sheet'!B2`), which is the node's child. */
  spillOf?: string
  /** For a range: the number of non-empty cells. */
  cells?: number
  /**
   * A cell with an error, or a formula stopped by a circular reference: the cell where the error originated
   * (`'Sheet'!A2`, or the node's own `ref` when it originated there), as in `CellIssueDetail.origin`. The path to it is
   * in the tree even beyond `depth`, and it passes through a range by way of the cell the error came from.
   */
  origin?: string
  children: ExplainNode[]
}

/** What `explain` returns: why a cell holds its value, as a tree of what it depends on and as text. */
export interface Explanation {
  root: ExplainNode
  nodes: number
  truncated: boolean
  /**
   * The tree as indented text (for logs, an LLM agent or a user) in the workbook's locale: in ru-RU its words, numbers,
   * formulas, booleans and errors are in Russian (`План!B2 = 480  ← ОКРУГЛ(B1*1,2;0)`), and a sheet name is quoted only
   * where Excel quotes it; the nodes keep the stored text. An error is written as the workbook's cells display it
   * (`#DIV/0!`), with the cell where it originated (`(error from S!A2)`).
   */
  text: string
}

/**
 * A workbook and its result: a session's, taken at a single point in time (`Session.getSnapshot`), or a workbook and
 * what `compute` returned for it, in the form `writeXlsx` accepts and `viewportOf` reads from a version a host saved.
 */
export interface BookSnapshot {
  book: Book
  result: ComputeResult
}
