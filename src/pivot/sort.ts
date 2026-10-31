/** A PivotTable's middle-pivot partition sort; equal values cross the partition on descending sorts alone. */
import { spend } from '../exec/budget.js'

type Bounds = { min: number; max: number; low: number; high: number }

export function sortPivotItems(items: number[], before: (pivot: number, item: number) => boolean, threads = 1, compare?: (a: number, b: number) => number): void {
  const swap = (a: number, b: number) => { const t = items[a]; items[a] = items[b]; items[b] = t }
  const bounds = (lo: number, hi: number): Bounds => {
    const b = { min: items[lo], max: items[lo], low: 1, high: 1 }
    for (let i = lo + 1; i <= hi; i++) {
      spend(1)
      const low = compare!(items[i], b.min), high = compare!(items[i], b.max)
      if (low < 0) { b.min = items[i]; b.low = 1 } else if (low === 0) b.low++
      if (high > 0) { b.max = items[i]; b.high = 1 } else if (high === 0) b.high++
    }
    return b
  }
  const sort = (lo: number, hi: number): void => {
    let span: Bounds | undefined
    while (lo < hi) {
      spend(1)
      if (compare && !span && hi - lo >= 64) span = bounds(lo, hi)
      swap(lo, Math.floor((lo + hi) / 2))
      const pivot = items[lo]
      // An extreme pivot puts every item on one side. Skip those comparisons without changing the swaps.
      if (span) {
        const low = compare!(pivot, span.min) === 0, high = compare!(pivot, span.max) === 0
        const equalBefore = before(pivot, pivot)
        if ((low && !equalBefore) || (high && equalBefore)) {
          if (equalBefore) swap(lo, hi--)
          else lo++
          if (low) span.low--
          if (high) span.high--
          if (!span.low || !span.high) span = undefined
          continue
        }
      }
      let last = lo
      for (let i = lo + 1; i <= hi; i++) {
        spend(1)
        if (before(pivot, items[i])) swap(++last, i)
      }
      span = undefined
      swap(lo, last)
      // Recurse into the smaller side, keeping the stack logarithmic even when every value is equal.
      if (last - lo <= hi - last) {
        sort(lo, last - 1)
        lo = last + 1
      } else {
        sort(last + 1, hi)
        hi = last - 1
      }
    }
  }
  const parts = items.length > 512 ? Math.min(threads, Math.floor(items.length / 256)) : 1
  if (parts <= 1) return sort(0, items.length - 1)
  const size = Math.ceil(items.length / parts)
  for (let lo = 0; lo < items.length; lo += size) sort(lo, Math.min(items.length - 1, lo + size - 1))
  const buffer = items.slice()
  for (let width = size; width < items.length; width *= 2) {
    for (let lo = 0; lo < items.length; lo += 2 * width) {
      const mid = Math.min(items.length, lo + width), end = Math.min(items.length, lo + 2 * width)
      let a = lo, b = mid
      for (let i = lo; i < end; i++) buffer[i] = b === end || (a < mid && !before(items[a], items[b])) ? items[a++] : items[b++]
    }
    for (let i = 0; i < items.length; i++) items[i] = buffer[i]
  }
}
