/**
 * Excel's sort order of values of any type, as SORT and SORTBY sort (`functions/lookup/order.ts`) and a host's display
 * filter does (`compareValues` of `xlcalc/text`): numbers, text, booleans, errors (equal), then blanks, last either way.
 */
import { collator } from './collate.js'

/** Sort rank of a type: numbers, text, booleans, errors, then blanks. */
const rank = (v: unknown) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : v === null ? 4 : 3)

/** The order of two values of one type, text by `text`. */
export const sameType =
  (text: (a: string, b: string) => number) =>
  <T>(x: T, y: T): number =>
    typeof x === 'string' ? text(x, y as string) : x < y ? -1 : x > y ? 1 : 0

const EN_US = /* @__PURE__ */ sameType(collator.compare)

/** The order of two values, one type by `same` (text by en-US's collation by default), errors equal; `desc` reverses it, blanks stay last. */
export function sortOrder<T>(a: T, b: T, desc: boolean, same: (a: T, b: T) => number = EN_US): number {
  const ra = rank(a)
  const rb = rank(b)
  if (ra === 4 || rb === 4) return ra - rb
  const c = ra !== rb ? ra - rb : ra === 3 ? 0 : same(a, b)
  return desc ? -c : c
}
