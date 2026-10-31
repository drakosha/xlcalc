/** A PivotTable's middle-pivot partition sort; equal values cross the partition on descending sorts alone. */
export function sortPivotItems(items: number[], before: (pivot: number, item: number) => boolean): void {
  const swap = (a: number, b: number) => { const t = items[a]; items[a] = items[b]; items[b] = t }
  const sort = (lo: number, hi: number): void => {
    while (lo < hi) {
      swap(lo, Math.floor((lo + hi) / 2))
      const pivot = items[lo]
      let last = lo
      for (let i = lo + 1; i <= hi; i++) if (before(pivot, items[i])) swap(++last, i)
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
  sort(0, items.length - 1)
}
