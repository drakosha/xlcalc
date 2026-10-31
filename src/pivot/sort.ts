/** A PivotTable's middle-pivot partition sort; equal values cross the partition on descending sorts alone. */
export function sortPivotItems(items: number[], before: (pivot: number, item: number) => boolean, threads = 1): void {
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
