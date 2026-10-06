/**
 * Information functions. Ported from formula.js 4.6.1 (src/information.js), Copyright (c) 2014 Sutoiku,
 * Inc., MIT License (THIRD_PARTY.md); TYPE is this engine's own.
 */
import type { Cx, FnDef } from '../args.js'
import { E, isErr, type Rng, type V } from '../values.js'
import type { ErrorCode } from '../contract.js'

/** ERROR.TYPE numbers; #ERROR! is how this engine reads #NULL! (1). #CYCLE! never reaches a formula. */
const ERROR_TYPE: Partial<Record<ErrorCode, number>> = { ERROR: 1, DIV0: 2, VALUE: 3, REF: 4, NAME: 5, NUM: 6, NA: 7, SPILL: 9 }

const typeOf = (v: V) => (typeof v === 'number' || v === null ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 4 : 16)

const is = (test: (v: V) => boolean): FnDef => ({ args: 'a', fn: test })

export const functions: Record<string, FnDef> = {
  'ERROR.TYPE': { args: 'a', fn: (v: V) => (isErr(v) ? (ERROR_TYPE[v.code] ?? E.NA) : E.NA) },
  ISBLANK: is((v) => v === null),
  ISERR: is((v) => isErr(v) && v.code !== 'NA'),
  ISERROR: is(isErr),
  // Excel: a boolean is #VALUE!, a fraction is truncated
  ISEVEN: { args: 'x', fn: (x: number) => Math.trunc(x) % 2 === 0 },
  ISLOGICAL: is((v) => typeof v === 'boolean'),
  ISNA: is((v) => isErr(v) && v.code === 'NA'),
  ISNONTEXT: is((v) => typeof v !== 'string'),
  ISNUMBER: is((v) => typeof v === 'number'),
  ISODD: { args: 'x', fn: (x: number) => Math.trunc(x) % 2 !== 0 },
  ISTEXT: is((v) => typeof v === 'string'),
  N: { args: 'a', fn: (v: V) => (typeof v === 'number' || isErr(v) ? v : v === true ? 1 : 0) },
  NA: { args: '', fn: () => E.NA },
  // a range of several cells is #VALUE! in a value position (type 16), as in Excel without dynamic arrays
  TYPE: { args: 'R', cx: true, fn: (cx: Cx, x: V | Rng) => typeOf(cx.scalar(x)) },
}
