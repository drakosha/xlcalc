/** ERROR.TYPE's numbers of the errors, as ERROR.TYPE answers them and SORT orders errors among themselves. */
import type { ErrorCode } from './contract.js'

/**
 * ERROR.TYPE numbers; #ERROR!, a formula that does not compute, is 1 as #NULL! is. #CYCLE! never reaches a formula, and
 * #BUSY! has no number Excel documents.
 */
export const ERROR_TYPE: Partial<Record<ErrorCode, number>> = { NULL: 1, ERROR: 1, DIV0: 2, VALUE: 3, REF: 4, NAME: 5, NUM: 6, NA: 7, GETTING_DATA: 8, SPILL: 9, CONNECT: 10, BLOCKED: 11, CALC: 14 }
