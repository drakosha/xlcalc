/**
 * The formula engine contract (subpath `./contract`): types only, no engine code. Any engine
 * implementation (currently formulas compiled to JS) sits behind this contract; the package's
 * behaviour tests are its acceptance suite.
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
  /** Display format, passed through untouched. */
  format?: string
  /** Arbitrary cell data (for example ids of source rows), passed through untouched. */
  meta?: JsonValue
}

export interface SheetInput {
  name: string
  /** Cells by A1 address (`B12`). */
  cells: Record<string, CellInput>
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
  sheets: SheetInput[]
  names?: NameInput[]
}

export type ErrorCode = 'DIV0' | 'NA' | 'NAME' | 'NUM' | 'REF' | 'VALUE' | 'CYCLE' | 'SPILL' | 'ERROR'

/** A computed value. Dates come out as serial numbers, numbers are rounded to 15 significant digits. */
export type CellValue = Scalar | { error: ErrorCode }

/** Formula check codes. `unsupported`: array arithmetic, or an external function with non-literal arguments. */
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

export interface ComputeOptions {
  /**
   * The clock for TODAY, NOW, RAND and RANDBETWEEN. By default the current moment and a random seed;
   * pass the `clock` of a result to reproduce that result exactly.
   */
  clock?: Clock
  externals?: Externals
}

export interface CellResult {
  value: CellValue
  /** Canonical formula text (sheet always quoted, function names upper-case). */
  formula?: string
  format?: string
  meta?: JsonValue
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
  /** sha256 of the canonical text of the values together with the engine configuration version. */
  hash: string
  /** sha256 of the input (sheets, cells, names; order does not matter). */
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

export interface CellChange extends CellRef {
  value: CellValue
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
 * recomputation instead.
 */
export type SessionOptions = ({ clock?: Clock; live?: false } | { live: true }) & { externals?: Externals }

/**
 * A session over a book: edits recompute only the cells that depend on them. The current clock of the
 * session (`clock`) is what a saved result should record: a full computation with it gives the same hash.
 */
export interface Session {
  readonly clock: Clock
  /** Applied atomically with one recomputation; returns the changed values. */
  setCells(edits: CellEdit[]): CellChange[]
  setNames(names: NameInput[]): CellChange[]
  /** Non-live sessions only. */
  setClock(clock: Clock): CellChange[]
  explain(sheet: string, cell: string, opts?: ExplainOptions): Explanation
  result(): ComputeResult
  close(): void
}

export interface FormulaEngine {
  compute(book: BookInput, opts?: ComputeOptions): ComputeResult
  validateFormula(text: string, ctx?: ValidateContext): ValidateResult
  validateName(name: string): NameCheck
  openSession(book: BookInput, opts?: SessionOptions): Session
}

export const MAX_FORMULA_LENGTH = 8_192
export const MAX_FORMULA_DEPTH = 64
