/**
 * RANK, RANK.EQ and RANK.AVG filled down a column over one long range: the range's numbers are sorted once and kept
 * with its values (range arrays are immutable and replaced on invalidation), and each call counts by binary search
 * what a pass over the list counts. The runtime takes this way before the general call (`runtime.ts`).
 */
import { coerce, type Arg, type Cx } from '../../args.js'
import { keptWith, VALUE_BYTES } from '../../range-cache.js'
import { lowerBound } from '../../sorted.js'
import { E, fitNumber, isErr, Rng, type CErr, type V } from '../../values.js'

/** The sizes `Cx.flat` keeps a range's values from, so they are the same array call after call. */
const KEPT_CELLS = 256
const sortedRanges = new WeakMap<readonly V[], Float64Array | CErr>()
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

/** The first position whose value is above `x`. */
function upperBound(xs: Float64Array, x: number): number {
  let lo = 0
  let hi = xs.length
  while (lo < hi) {
    const mid = (lo + hi) >> 1
    if (xs[mid] <= x) lo = mid + 1
    else hi = mid
  }
  return lo
}

/**
 * RANK of a number over one range of 256 cells or more and an order that is a number, a boolean, a blank or not given;
 * undefined where the call goes the general way, which answers the same (any error or text among the other arguments
 * decides the order of the answers there).
 */
export function rankInRange(cx: Cx, average: boolean, args: readonly (Arg | undefined)[]): V | undefined {
  const [x, list, order] = args
  if (args.length > 3 || !(list instanceof Rng)) return undefined
  const b = cx.bounds(list)
  if (b.h * b.w < KEPT_CELLS) return undefined
  const n = coerce('n', x!, cx)
  if (typeof n !== 'number') return undefined
  const o = order === undefined ? undefined : cx.scalar(order)
  if (o !== undefined && o !== null && typeof o !== 'number' && typeof o !== 'boolean') return undefined
  const values = cx.flat(list)
  const desc = !(o ?? 0)
  let sorted = sortedRanges.get(values)
  if (!sorted) {
    if (!readOnce.has(values)) {
      readOnce.add(values)
      return passRank(values, n, desc, average)
    }
    sortedRanges.set(values, (sorted = sortedOf(values)))
    keptWith(values, 64 + (isErr(sorted) ? 0 : sorted.length * VALUE_BYTES))
  }
  if (isErr(sorted)) return sorted
  const lo = lowerBound(sorted, n)
  const hi = upperBound(sorted, n)
  return answer(desc ? sorted.length - hi : lo, hi - lo, average)
}

/** The rank from the numbers before `x` in the order and those equal to it; none equal is #N/A. */
const answer = (before: number, equal: number, average: boolean): V => (equal ? fitNumber(average ? before + (equal + 1) / 2 : before + 1) : E.NA)

/** The rank by one pass over the range's values, as RANK counts its list. */
function passRank(values: readonly V[], x: number, desc: boolean, average: boolean): V {
  let before = 0
  let equal = 0
  for (const v of values) {
    if (typeof v === 'number') {
      if (v === x) equal++
      else if (desc ? v > x : v < x) before++
    } else if (isErr(v)) return v
  }
  return answer(before, equal, average)
}
