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

/** The bands between the cuts: the sorted cuts made unique, each from one cut to before the next. */
export function bands(cuts: Iterable<number>): [number, number][] {
  const xs = [...new Set(cuts)].sort((a, b) => a - b)
  return xs.slice(0, -1).map((x, i) => [x, xs[i + 1] - 1])
}
