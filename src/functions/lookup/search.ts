/** How the lookups compare, as in Excel: within a type, text in any case, blanks and errors in the data never match. */
import { wildRe } from '../../criteria.js'
import { compareValues as compare, isErr, sameType, type CErr, type V } from '../../values.js'

export type Key = Exclude<V, null | CErr>

/** The lookup value: an error is the answer, a blank is looked up as 0. */
export const keyOf = (key: V): Key | CErr => (key === null ? 0 : key)

const comparable = (v: V, k: Key) => v !== null && !isErr(v) && sameType(v, k)

/** Exact match; with `wild`, `*`, `?` and `~` in text work as in the criteria. */
function exact(k: Key, wild: boolean): (v: V) => boolean {
  const re = wild && typeof k === 'string' ? wildRe(k) : null
  if (re) return (v) => typeof v === 'string' && re.test(v)
  return (v) => comparable(v, k) && compare(v, k) === 0
}

/**
 * Excel's approximate match is a binary search: the last position not greater than the key (`desc`: not
 * less), or −1; a midpoint of another type, blank or error moves to the nearest comparable value below (else above).
 */
export function bisect(at: (i: number) => V, n: number, k: Key, desc = false): number {
  let lo = 0
  let hi = n - 1
  let found = -1
  while (lo <= hi) {
    const mid = (lo + hi) >> 1
    let m = mid
    while (m >= lo && !comparable(at(m), k)) m--
    if (m < lo) {
      m = mid + 1
      while (m <= hi && !comparable(at(m), k)) m++
      if (m > hi) break
    }
    if (compare(at(m), k) * (desc ? -1 : 1) <= 0) {
      found = m
      lo = m + 1
    } else hi = m - 1
  }
  return found
}

/** Position among n values: exact (the first match) or approximate (mode 1 ascending, −1 descending), or −1. */
export function find(at: (i: number) => V, n: number, k: Key, mode: number, wild: boolean): number {
  if (mode !== 0) return bisect(at, n, k, mode < 0)
  const test = exact(k, wild)
  for (let i = 0; i < n; i++) if (test(at(i))) return i
  return -1
}
