/**
 * How the lookups compare, as in Excel: within a type, text in any case by the book's locale (`Compare`,
 * `Rules.compare`), blanks and errors in the data never match.
 */
import { isArea } from '../../arrays/lift.js'
import { wildRe } from '../../criteria.js'
import { E, isErr, Rng, sameType, type CErr, type V } from '../../values.js'
import type { Cx } from '../../args.js'
import { indexedExact, scannedExact } from './exact.js'

export type Key = Exclude<V, null | CErr>

/** Excel's order of two values that are not errors. */
export type Compare = (a: V, b: V) => number

/** The lookup value: an error is the answer, a blank is looked up as 0. */
export const keyOf = (key: V): Key | CErr => (key === null ? 0 : key)

/** A value of the data that can match the key: of its type, no blank and no error. */
export const comparable = (v: V, k: Key) => v !== null && !isErr(v) && sameType(v, k)

/** Exact match; with `wild`, `*`, `?` and `~` work in text, `~` escaping any character. */
export function exactMatch(line: Line, k: Key, wild: boolean, compare: Compare, last = false): number {
  const re = wild && typeof k === 'string' ? wildRe(k, true) : null
  if (!re && line.values) {
    const hit = indexedExact(line.values, line.stride ?? 1, line.filled, k, compare, last)
    if (hit !== undefined) return hit
  }
  for (let j = 0; j < line.filled; j++) {
    const i = last ? line.filled - 1 - j : j
    const v = line.at(i)
    if (re ? typeof v === 'string' && re.test(v) : comparable(v, k) && compare(v, k) === 0) {
      if (!re && line.values) scannedExact(line.values, line.stride ?? 1, line.filled, k, compare, j + 1)
      return i
    }
  }
  if (!re && line.values) scannedExact(line.values, line.stride ?? 1, line.filled, k, compare, line.filled)
  return -1
}

/** Values along a range: `n` positions as referenced, of which the first `filled` can hold cells (the rest are blank). */
export interface Line {
  at: (i: number) => V
  n: number
  filled: number
  /** Immutable storage for cached exact searches, starting at its first value. */
  values?: readonly V[]
  stride?: number
}

/**
 * Excel's binary search over the referenced positions (fitted to its answers on unsorted, mixed and whole-column
 * data): past the last non-blank cell it goes back, a blank or another type before it is passed over for the
 * next comparable value, an equal value ends at the last of its run. The last position not greater than k, or −1.
 */
export function bisect({ at, n, filled }: Line, k: Key, compare: Compare): number {
  let last = filled - 1
  while (last >= 0 && at(last) === null) last--
  let lo = 0
  let hi = n - 1
  let found = -1
  while (lo <= hi) {
    const mid = (lo + hi) >> 1
    if (mid > last) {
      hi = mid - 1
      continue
    }
    let m = mid
    while (m <= hi && !comparable(at(m), k)) m++
    if (m > hi) for (m = mid - 1; m >= lo && !comparable(at(m), k); m--);
    if (m < lo) break
    const c = compare(at(m), k)
    if (c === 0) {
      while (m + 1 <= last && comparable(at(m + 1), k) && compare(at(m + 1), k) === 0) m++
      return m
    }
    if (c < 0) {
      found = m
      lo = m + 1
    } else hi = m - 1
  }
  return found
}

/** Excel's match on descending data scans: the first equal value, else the last value before one that is smaller. */
function descend({ at, filled }: Line, k: Key, compare: Compare): number {
  let found = -1
  for (let i = 0; i < filled; i++) {
    const v = at(i)
    if (!comparable(v, k)) continue
    const c = compare(v, k)
    if (c === 0) return i
    if (c < 0) break
    found = i
  }
  return found
}

/** Position on a line: exact (the first match), approximate ascending (1) or descending (−1), or −1. */
export function find(line: Line, k: Key, mode: number, wild: boolean, compare: Compare): number {
  if (mode > 0) return bisect(line, k, compare)
  if (mode < 0) return descend(line, k, compare)
  return exactMatch(line, k, wild, compare)
}

/**
 * A column index or match type: a cell that is no number gives #REF!, a value that is none its own error, as
 * in Excel 365; a range of several cells is #REF! for a column index and its top-left cell for a match type.
 */
export function indexArg(cx: Cx, x: V | Rng, range: 'index' | 'type' = 'index'): number | CErr {
  if (!(x instanceof Rng)) return cx.rules.toNum(x)
  if (isArea(x)) {
    if (range === 'index') return E.REF
    x = new Rng(x.s, x.r1, x.c1, x.r1, x.c1)
  }
  const n = cx.rules.toNum(cx.scalar(x))
  return isErr(n) ? E.REF : n
}
