/** Reusable row-major prefixes for single-range SUM/COUNT/MIN/MAX, with unrounded sums and first-error order. */
import type { Bounds } from './args.js'
import { spend } from './exec/budget.js'
import { RangeCache } from './range-cache.js'
import type { Area } from './spill/state.js'
import { fitNumber, isErr, Rng, type CErr, type V } from './values.js'

export type AggregateName = 'SUM' | 'COUNT' | 'MIN' | 'MAX'
interface Partial {
  sum: number
  count: number
  min: number
  max: number
  error?: CErr
}
const CHUNK = 256

export class RangeAggregates {
  private readonly cache = new RangeCache<Partial>()

  invalidate(area: Area): void { this.cache.invalidate(area) }

  read(name: AggregateName, range: Rng, bounds: () => Bounds, cell: (s: number, r: number, c: number) => V): V {
    let result = this.cache.get(range.key)
    if (!result) {
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
        if (previous) { result = { ...previous }; at = end }
        else {
          // Checkpoints also support descending evaluation order, where the longest range is read first.
          for (let p = start + Math.floor((end - start) / block) * block - 1; p >= start; p -= block) {
            const hit = this.cache.get(prefix(p).key)
            if (hit) { result = { ...hit }; at = p + 1; break }
          }
        }
      }
      result ??= { sum: 0, count: 0, min: Infinity, max: -Infinity }
      const take = (value: V) => {
        if (typeof value === 'number') {
          result!.sum += value
          result!.count++
          if (value < result!.min) result!.min = value
          if (value > result!.max) result!.max = value
        } else if (isErr(value)) result!.error ??= value
      }
      for (; at <= end; at++) {
        if (horizontal) { spend(1); take(cell(range.s, range.r1, at)) }
        else {
          spend(b.w)
          for (let c = b.c1; c <= b.c2; c++) take(cell(range.s, at, c))
        }
        if ((at - start + 1) % block === 0) this.cache.set(prefix(at).key, { ...result })
      }
      if (b.h * b.w >= CHUNK) this.cache.set(range.key, result)
    }
    if (name === 'COUNT') return result.count
    if (result.error) return result.error
    return fitNumber(name === 'SUM' ? result.sum : result.count === 0 ? 0 : name === 'MIN' ? result.min : result.max)
  }
}
