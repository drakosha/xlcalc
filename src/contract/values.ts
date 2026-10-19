/** Values and places: a plain value and JSON, a computed value and its error, the check codes, the clock, a cell and a range of a book, a sheet's positions and rectangles numbered from 1, and a refusal. */

/** A plain value: a number, text, a boolean, or null for an empty cell. */
export type Scalar = number | string | boolean | null

/** Data a host keeps on a cell (`CellInput.meta`), as JSON. */
export type JsonValue = Scalar | JsonValue[] | { [key: string]: JsonValue }

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

/** A cell of a book: its sheet's name and its A1 address (`B2`). */
export interface CellRef {
  sheet: string
  cell: string
}

/** A cell's row and column, numbered from 1 as A1 numbers them: `B12` is row 12, column 2. */
export interface CellPosition {
  row: number
  col: number
}

/** A rectangle of a sheet: rows and columns numbered from 1, bounds inclusive; a whole column runs from row 1 to `MAX_ROWS`. */
export interface CellRect {
  top: number
  left: number
  bottom: number
  right: number
}

/** A range's rectangle and what its text names: cells (`B2:D9`, `B2`), whole columns (`B:D`) or whole rows (`2:9`). */
export interface RangeRect extends CellRect {
  kind: 'cells' | 'columns' | 'rows'
}

/** The lines a key of `SheetInput.columns` or `rows` names (`C`, `B:D`, `5`, `5:9`), numbered from 1. */
export interface LineSpan {
  first: number
  last: number
}

/**
 * A check or an edit that did not pass: a code a program reads, a message for people (in the book's locale where the
 * call has one, English otherwise) and, for typed text, the position in it. Every check answers `{ ok: true, … }` or
 * a `Refusal` with its own codes.
 */
export interface Refusal<Code extends string = string> {
  ok: false
  code: Code
  message: string
  at?: number
}

/** A rectangle on a sheet: `B2`, `B2:D9`, whole columns `B:D` or whole rows `2:9` (no `$`, no sheet). */
export interface RangeRef {
  sheet: string
  range: string
}
