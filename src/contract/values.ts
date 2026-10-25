/** Values and places: a plain value and JSON, a computed value and its error, the check codes, the clock, a cell and a range of a book, a sheet's positions and rectangles numbered from 1, and a refusal. */

/** A plain value: a number, text, a boolean, or null for an empty cell. */
export type Scalar = number | string | boolean | null

/** Data a host keeps on a cell (`CellInput.meta`), as JSON. */
export type JsonValue = Scalar | JsonValue[] | { [key: string]: JsonValue }

/**
 * `CALC`: Excel's #CALC!, an empty array (`FILTER` that matches nothing) or text longer than 32,767 characters
 * (`CONCAT`). `SPILL`: an array that cannot spill from its formula's cell (the range is not empty, overlaps a merged
 * area or runs off the sheet). `NULL`: Excel's #NULL!, ranges that do not intersect under the intersection operator
 * (`A1:A3 C1:C3`). `ERROR`: a formula that cannot be calculated (for example, one that does not parse).
 */
export type ErrorCode = 'DIV0' | 'NA' | 'NAME' | 'NUM' | 'REF' | 'VALUE' | 'NULL' | 'CYCLE' | 'SPILL' | 'CALC' | 'ERROR'

/** A calculated value. Dates are returned as serial numbers, and numbers are rounded to 15 significant digits. */
export type CellValue = Scalar | { error: ErrorCode }

/**
 * Formula check codes. `unsupported`: an external function called with arguments that are not literals.
 * `unknown_name` also covers tables: an unknown table or column, or a structured reference without a table name outside
 * a table.
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

/**
 * The point in time and the random seed from which TODAY, NOW, RAND and RANDBETWEEN are calculated: ISO 8601 with Z or
 * an offset (`2026-10-04T12:00:00Z`), or a date alone, read as midnight UTC.
 */
export interface Clock {
  now: string
  seed: number
}

/** A cell of a workbook: its sheet's name and its A1 address (`B2`). */
export interface CellRef {
  sheet: string
  cell: string
}

/** A cell's row and column, numbered from 1 as in A1 notation: `B12` is row 12, column 2. */
export interface CellPosition {
  row: number
  col: number
}

/** A rectangle on a sheet: rows and columns numbered from 1, bounds inclusive; an entire column runs from row 1 to `MAX_ROWS`. */
export interface CellRect {
  top: number
  left: number
  bottom: number
  right: number
}

/** A range's rectangle and what its text refers to: cells (`B2:D9`, `B2`), entire columns (`B:D`) or entire rows (`2:9`). */
export interface RangeRect extends CellRect {
  kind: 'cells' | 'columns' | 'rows'
}

/** The columns or rows a key of `SheetInput.columns` or `rows` refers to (`C`, `B:D`, `5`, `5:9`), numbered from 1. */
export interface LineSpan {
  first: number
  last: number
}

/**
 * A check or an edit that failed: a code for programs, a message for users (in the workbook's locale when the call has
 * one, otherwise in English) and, for typed text, the position in that text. Every check returns either
 * `{ ok: true, … }` or a `Refusal` with its own codes.
 */
export interface Refusal<Code extends string = string> {
  ok: false
  code: Code
  message: string
  at?: number
}

/** A rectangle on a sheet: `B2`, `B2:D9`, entire columns `B:D` or entire rows `2:9` (no `$`, no sheet name). */
export interface RangeRef {
  sheet: string
  range: string
}
