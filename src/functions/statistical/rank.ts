/**
 * RANK, RANK.EQ and RANK.AVG: the count of numbers before x in the order and of those equal to it. Filled down a column
 * over one long range, the range's numbers are sorted once and kept with its values (range arrays are immutable and
 * replaced on invalidation), and each call counts by binary search; the runtime takes this way before the general call
 * (`runtime.ts`).
 */
import type { Arg, Cx } from '../../args.js'
import { keepAlong, KEPT_RANGE } from '../../range-cache.js'
import { lowerBound, upperBound } from '../../sorted.js'
import { E, fitNumber, isErr, Rng, type CErr, type V } from '../../values.js'

/** RANK's order: a number or a boolean; text, even "0", is #VALUE!. */
function descending(order: V): boolean | V {
  if (typeof order === 'string') return E.VALUE
  return !order
}

/** The rank from the numbers before x in the order and those equal to it; none equal is #N/A. */
const rankOf = (before: number, equal: number, average: boolean): number | CErr => (equal ? (average ? before + (equal + 1) / 2 : before + 1) : E.NA)

/** The rank by one pass over values (numbers counted, the first error the answer, the rest skipped). */
function counted(values: readonly V[], x: number, desc: boolean, average: boolean): number | CErr {
  let before = 0
  let equal = 0
  for (const v of values) {
    if (typeof v === 'number') {
      if (v === x) equal++
      else if (desc ? v > x : v < x) before++
    } else if (isErr(v)) return v
  }
  return rankOf(before, equal, average)
}

function rank(x: number, xs: number[], order: V | undefined, average: boolean): V {
  const desc = descending(order ?? 0)
  return typeof desc === 'boolean' ? counted(xs, x, desc, average) : desc
}

export const rankEq = (x: number, xs: number[], order?: V) => rank(x, xs, order, false)
export const rankAvg = (x: number, xs: number[], order?: V) => rank(x, xs, order, true)

const sortedRanges = new WeakMap<readonly unknown[], Float64Array | CErr>()
/** Values read once: sorted on the second read, which shows the range keeps them (a full budget does not). */
const readOnce = new WeakSet<readonly V[]>()

/** The numbers of a range ascending (text, booleans and blanks skipped), or its first error, as the list kind reads it. */
function sortedOf(values: readonly V[]): Float64Array | CErr {
  const out: number[] = []
  for (const v of values)
    if (typeof v === 'number') out.push(v)
    else if (isErr(v)) return v
  return Float64Array.from(out).sort()
}

/**
 * RANK of a number `x` (undefined where it is an error) over one range of 256 cells or more and an order that is a
 * number, a boolean, a blank or not given; undefined where the call goes the general way, which answers the same (any
 * error or text among the other arguments decides the order of the answers there).
 */
export function rankInRange(cx: Cx, average: boolean, x: number | undefined, list: Arg | undefined, order: Arg | undefined): V | undefined {
  if (x === undefined || !(list instanceof Rng)) return undefined
  const b = cx.bounds(list)
  if (b.h * b.w < KEPT_RANGE) return undefined
  const o = order === undefined ? undefined : cx.scalar(order)
  if (o !== undefined && o !== null && typeof o !== 'number' && typeof o !== 'boolean') return undefined
  const desc = !(o ?? 0)
  const values = cx.flat(list)
  let sorted = sortedRanges.get(values)
  if (!sorted) {
    if (!readOnce.has(values)) {
      readOnce.add(values)
      return fit(counted(values, x, desc, average))
    }
    const made = sortedOf(values)
    sorted = keepAlong(sortedRanges, values, made, isErr(made) ? 0 : made.length)
  }
  if (isErr(sorted)) return sorted
  const lo = lowerBound(sorted, x)
  const hi = upperBound(sorted, x)
  return fit(rankOf(desc ? sorted.length - hi : lo, hi - lo, average))
}

/** A number answered as the general call answers it (`fitNumber`). */
const fit = (r: number | CErr): V => (isErr(r) ? r : fitNumber(r))
