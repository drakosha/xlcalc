/**
 * Excel's sort order of values of any type, as SORT and SORTBY sort (`functions/lookup/order.ts`) and the Sort command
 * of a host's display filter does (`compareValues` of `xlcalc/text`): numbers, text, booleans, errors, then blanks, last
 * either way. SORT orders errors by ERROR.TYPE's number; the Sort command takes them all as equal.
 */
import { collator } from './collate.js'
import type { ErrorCode } from './contract.js'
import { ERROR_TYPE } from './error-type.js'

/** An error's ERROR.TYPE number, an engine's error (`code`) or a cell's (`error`) alike. */
const errorType = (v: unknown) => {
  const e = v as { code?: ErrorCode; error?: ErrorCode }
  return ERROR_TYPE[(e.code ?? e.error)!] ?? 0
}

/** Sort rank of a type: numbers, text, booleans, errors, then blanks. */
const rank = (v: unknown) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : v === null ? 4 : 3)

/** The order of two values of one type, text by `text`. */
export const sameType =
  (text: (a: string, b: string) => number) =>
  <T>(x: T, y: T): number =>
    typeof x === 'string' ? text(x, y as string) : x < y ? -1 : x > y ? 1 : 0

const EN_US = /* @__PURE__ */ sameType(collator.compare)

/**
 * The order of two values, one type by `same` (text by en-US's collation by default), errors equal as the Sort command
 * takes them or, with `byErrorType`, by ERROR.TYPE's number (#DIV/0! before #NAME?, as SORT orders them); `desc`
 * reverses it, blanks stay last.
 */
export function sortOrder<T>(a: T, b: T, desc: boolean, same: (a: T, b: T) => number = EN_US, byErrorType = false): number {
  const ra = rank(a)
  const rb = rank(b)
  if (ra === 4 || rb === 4) return ra - rb
  const c = ra !== rb ? ra - rb : ra === 3 ? (byErrorType ? errorType(a) - errorType(b) : 0) : same(a, b)
  return desc ? -c : c
}
