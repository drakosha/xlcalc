/** The formula AST, the parser's error with its check code, and the text of error values. */
import type { CheckCode, ErrorCode } from '../contract.js'
import type { TableSpec } from './table.js'

export interface RangeRef {
  /** Zero-based rows and columns; a whole column has r1 = 0, r2 = MAX_ROWS - 1. */
  r1: number
  c1: number
  r2: number
  c2: number
  kind: 'cell' | 'range' | 'cols' | 'rows'
  /** Source text of the reference without the sheet (`$D$2:$D$3401`), for the canonical form. */
  text: string
  /** A spill reference (`E2#`, a cell): the range the cell's array spilled into. */
  spill?: true
}

export type Ast =
  | { k: 'num'; v: number }
  | { k: 'str'; v: string }
  | { k: 'bool'; v: boolean }
  /** `sheet` and `ref` keep what Excel writes around a reference it lost: `'S'!#REF!`, `#REF!A1`, `#REF!#REF!`, `#REF!#` (a spill reference). */
  | { k: 'err'; code: ErrorCode; sheet?: string; ref?: string }
  | { k: 'ref'; sheet: string | null; ref: RangeRef }
  /** A reference across the sheets from `first` to `last` in the book's order (`'Sheet1:Sheet3'!A1:B2`). */
  | { k: 'ref3d'; first: string; last: string; ref: RangeRef }
  /**
   * A structured reference: a table's areas and columns (`Sales[[#Headers],[Qty]]`); `table` null inside a table
   * (`[@Qty]`), whose name the formula's cell gives.
   */
  | { k: 'table'; table: string | null; spec: TableSpec }
  /** A defined name; with `sheet`, the name as that sheet sees it (`'Sheet2'!rate`). */
  | { k: 'name'; name: string; sheet?: string }
  /** Implicit intersection (`@x`): a reference's cell in the formula's row or column, an array's top-left value. */
  | { k: 'at'; a: Ast }
  /** A union of references (`(A1:A3,C1:C3)`, or a name's formula `'S'!A1,'S'!C1`). */
  | { k: 'union'; items: Ast[] }
  | { k: 'fn'; name: string; args: Ast[] }
  | { k: 'un'; op: '-' | '+'; a: Ast }
  | { k: 'pct'; a: Ast }
  | { k: 'bin'; op: BinOp; a: Ast; b: Ast }
  | { k: 'paren'; a: Ast }
  /** An array constant `{1,2;3,4}`: rows of numbers, text, booleans and errors. */
  | { k: 'arr'; rows: Lit[][] }
  /** An argument left out between commas (`IF(A1,,1)`). */
  | { k: 'missing' }
  /**
   * LET: each name bound to its value in turn, then the calculation; a call that ends with a value (an even
   * number of arguments, which Excel loads) answers that value. A name's `slot` numbers it among the
   * formula's LET names.
   */
  | { k: 'let'; binds: Bind[]; body?: Ast }
  /** A LET name where it is seen: the binding `slot` of the formula. */
  | { k: 'local'; name: string; slot: number }

export interface Bind {
  name: string
  slot: number
  value: Ast
}

/** A value of an array constant. */
export type Lit = Extract<Ast, { k: 'num' | 'str' | 'bool' | 'err' }>

export type BinOp = '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '<' | '>' | '<=' | '>='

export class FormulaError extends Error {
  constructor(
    readonly code: CheckCode,
    message: string,
    readonly at: number,
  ) {
    super(message)
  }
}

export function bad(at: number, message: string, code: CheckCode = 'syntax'): never {
  throw new FormulaError(code, message, at)
}

export const ERROR_TEXT: Record<ErrorCode, string> = {
  DIV0: '#DIV/0!',
  NA: '#N/A',
  NAME: '#NAME?',
  NUM: '#NUM!',
  REF: '#REF!',
  VALUE: '#VALUE!',
  CYCLE: '#CYCLE!',
  SPILL: '#SPILL!',
  CALC: '#CALC!',
  ERROR: '#ERROR!',
}
