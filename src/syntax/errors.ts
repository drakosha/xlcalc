/**
 * Error codes as Excel writes them in formulas and cells in en-US (`#DIV/0!`), and the sets of error texts each reader
 * takes; data alone, for the locale files too.
 */
import type { ErrorCode } from '../contract.js'

export const ERROR_TEXT: Record<ErrorCode, string> = {
  DIV0: '#DIV/0!',
  NA: '#N/A',
  NAME: '#NAME?',
  NUM: '#NUM!',
  REF: '#REF!',
  VALUE: '#VALUE!',
  NULL: '#NULL!',
  CYCLE: '#CYCLE!',
  SPILL: '#SPILL!',
  CALC: '#CALC!',
  ERROR: '#ERROR!',
}

/** Error codes by their text. */
export const ERROR_CODE: Readonly<Record<string, ErrorCode>> = Object.fromEntries(Object.entries(ERROR_TEXT).map(([code, text]) => [text, code as ErrorCode]))

const texts = (codes: readonly ErrorCode[]) => codes.map((c) => ERROR_TEXT[c])

/** The errors text typed into a cell is, by the locale's words (the entry set `typed`: `#SPILL!` and newer are text), in ERROR.TYPE's order. */
export const TYPED_ERRORS: readonly string[] = texts(['NULL', 'DIV0', 'VALUE', 'REF', 'NAME', 'NUM', 'NA'])

/**
 * The error literals a formula reads: the typed ones and `#SPILL!`, the formula of a file's cell valued so. Not `#CALC!`
 * (a file's cell valued so is left out; Excel loads neither `=#CALC!` nor `=#SPILL!`: the ru-RU snapshot's `rejected`),
 * nor the engine's own `#CYCLE!` and `#ERROR!`.
 */
export const FORMULA_ERRORS: readonly string[] = [...TYPED_ERRORS, ERROR_TEXT.SPILL]

/** The errors a file's cell may hold as its value or a formula's that the model has a code for: a formula's and `#CALC!`. */
export const CELL_ERRORS: readonly string[] = [...FORMULA_ERRORS, ERROR_TEXT.CALC]

/**
 * Excel's error literals in invariant spelling, which a locale may translate (`Locale.errors`): the model's but its own
 * `#CYCLE!` and `#ERROR!`, and those it has no code for; in the order a locale's check lists them.
 */
export const LOCALE_ERRORS: readonly string[] = [...TYPED_ERRORS, '#GETTING_DATA', ERROR_TEXT.SPILL, ERROR_TEXT.CALC, '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!', '#PYTHON!']

/** An error's text in a locale's words (`errors`, by stored text: `#DIV/0!` is `#ДЕЛ/0!` in ru-RU), the stored text where it has none. */
export const errorIn = (code: ErrorCode, errors?: Readonly<Record<string, string>>): string => errors?.[ERROR_TEXT[code]] ?? ERROR_TEXT[code]
