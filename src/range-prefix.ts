/** Reusable row-major prefixes. The fold owns numeric semantics; this class owns geometry, checkpoints and invalidation. */
import type { Bounds } from './args.js'
import { spend } from './exec/budget.js'
import { RangeCache, type CacheBudget } from './range-cache.js'
import type { Area } from './spill/state.js'
import { Rng, type V } from './values.js'

interface Fold<T> {
  create(): T
  copy(value: T): T
  bytes: number
}
const CHUNK = 256

export class RangePrefixes<T> {
  private readonly cache: RangeCache<T>
  constructor(private readonly fold: Fold<T>, budget: CacheBudget) { this.cache = new RangeCache(budget) }

  invalidate(area: Area): void { this.cache.invalidate(area) }
  clear(): void { this.cache.clear() }

  read(range: Rng, bounds: () => Bounds, cell: (s: number, r: number, c: number) => V, add: (value: T, cell: V) => void): T {
    let result = this.cache.get(range.key)
    if (result !== undefined) return result
    const { fold } = this
    const b = bounds()
    const horizontal = range.r1 === range.r2
    const start = horizontal ? range.c1 : range.r1
    const end = horizontal ? b.c2 : b.r2
    const block = horizontal ? CHUNK : Math.max(1, Math.floor(CHUNK / b.w))
    const prefix = (at: number) => horizontal
      ? new Rng(range.s, range.r1, range.c1, range.r2, at)
      : new Rng(range.s, range.r1, range.c1, at, range.c2)
    let at = start
    if (end > start) {
      const previous = this.cache.get(prefix(end - 1).key)
      if (previous !== undefined) { result = fold.copy(previous); at = end }
      else {
        // Descending evaluation can reuse checkpoints made by the longest range's first read.
        for (let p = start + Math.floor((end - start) / block) * block - 1; p >= start; p -= block) {
          const hit = this.cache.get(prefix(p).key)
          if (hit !== undefined) { result = fold.copy(hit); at = p + 1; break }
        }
      }
    }
    result ??= fold.create()
    for (; at <= end; at++) {
      if (horizontal) { spend(1); add(result, cell(range.s, range.r1, at)) }
      else {
        spend(b.w)
        for (let c = b.c1; c <= b.c2; c++) add(result, cell(range.s, at, c))
      }
      if ((at - start + 1) % block === 0) this.cache.set(prefix(at).key, fold.copy(result), fold.bytes)
    }
    if (b.h * b.w >= CHUNK) this.cache.set(range.key, result, fold.bytes)
    return result
  }
}
