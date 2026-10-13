/** First element whose numeric key is at least `value`, or the sorted array's length. */
export function lowerBound<T>(xs: ArrayLike<T>, value: number, key: (x: T) => number): number {
  let lo = 0
  let hi = xs.length
  while (lo < hi) {
    const mid = (lo + hi) >> 1
    if (key(xs[mid]) < value) lo = mid + 1
    else hi = mid
  }
  return lo
}
