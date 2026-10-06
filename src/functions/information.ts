/**
 * Information functions.
 *
 * Ported from formula.js 4.6.1 (https://github.com/formulajs/formulajs), src/information.js: Copyright (c)
 * 2014 Sutoiku, Inc., MIT License (full text in THIRD_PARTY.md). Changed: the engine's error values instead
 * of formula.js error objects; ERROR.TYPE knows #SPILL! (9).
 */
import { E, isErr, type V } from '../values.js'
import type { FnDef } from '../args.js'
import type { ErrorCode } from '../contract.js'

/** ERROR.TYPE numbers; #ERROR! is how this engine reads #NULL! (1). #CYCLE! never reaches a formula. */
const ERROR_TYPE: Partial<Record<ErrorCode, number>> = { ERROR: 1, DIV0: 2, VALUE: 3, REF: 4, NAME: 5, NUM: 6, NA: 7, SPILL: 9 }

const is = (test: (v: V) => boolean): FnDef => ({ args: 'a', fn: test })

export const functions: Record<string, FnDef> = {
  'ERROR.TYPE': { args: 'a', fn: (v: V) => (isErr(v) ? (ERROR_TYPE[v.code] ?? E.NA) : E.NA) },
  ISBLANK: is((v) => v === null),
  ISERR: is((v) => isErr(v) && v.code !== 'NA'),
  ISERROR: is(isErr),
  ISLOGICAL: is((v) => typeof v === 'boolean'),
  ISNA: is((v) => isErr(v) && v.code === 'NA'),
  ISNONTEXT: is((v) => typeof v !== 'string'),
  ISNUMBER: is((v) => typeof v === 'number'),
  ISTEXT: is((v) => typeof v === 'string'),
  N: { args: 'a', fn: (v: V) => (typeof v === 'number' || isErr(v) ? v : v === true ? 1 : 0) },
  NA: { args: '', fn: () => E.NA },
}
