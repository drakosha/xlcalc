/**
 * The formula engine contract (subpath `./contract`): types only, no engine code. Any engine
 * implementation (currently formulas compiled to closures or to JS text) sits behind this contract; the
 * package's behaviour tests are its acceptance suite.
 *
 * A formula is stored in the `formula` field without the leading `=` (one given is dropped), separately from the value:
 * a string from data never becomes a formula. Function names are English, arguments are separated
 * by commas, the decimal separator is a dot.
 */
export type { Locale } from './locale/types.js'
export type * from './contract/values.js'
export type * from './contract/book.js'
export type * from './contract/looks.js'
export type * from './contract/results.js'
export type * from './contract/text.js'
export type * from './contract/edits.js'
export type * from './contract/typing.js'
export type * from './contract/structure.js'
export type * from './contract/reading.js'
export type * from './contract/session.js'

/** Excel's limit on the length of a formula's text: 8,192 characters. */
export const MAX_FORMULA_LENGTH = 8_192
/** Excel's limit on nested function calls in a formula: 64. */
export const MAX_FORMULA_DEPTH = 64
