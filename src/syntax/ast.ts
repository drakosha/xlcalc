/** The formula AST, the parser's error with its check code, and the text of error values. */
import type { CheckCode, ErrorCode } from '../contract.js'

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
  | { k: 'name'; name: string }
  | { k: 'fn'; name: string; args: Ast[] }
  | { k: 'un'; op: '-' | '+'; a: Ast }
  | { k: 'pct'; a: Ast }
  | { k: 'bin'; op: BinOp; a: Ast; b: Ast }
  | { k: 'paren'; a: Ast }
  /** An array constant `{1,2;3,4}`: rows of numbers, text, booleans and errors. */
  | { k: 'arr'; rows: Lit[][] }

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
