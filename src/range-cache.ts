/**
 * Cached ranges indexed by their referenced bounds, including their unoccupied cells. The caches of range values and
 * criteria indexes keep a value from a range's second read on, and count their bytes in a `CacheBudget` shared by the
 * books of one heap (`docs/cache.md`): past its cap the least recently read entries are dropped between formulas, so
 * a formula's ranges keep their shape while it computes.
 */
import { AreaIndex } from './area-index.js'
import type { Area } from './spill/state.js'

interface Entry<T> extends Area {
  key: string
  /** None for a mark (`admits`). */
  value: T | undefined
  /** Bytes by the budget's estimate, the entry's own with its value's. */
  size: number
  /** The budget's read count when it was last read. */
  used: number
}

/** Bytes of an entry besides its value, a mark's too: its key, its bounds and its places in the maps, as measured on Node. */
const ENTRY_BYTES = 256

/** Bytes the range caches of the books in one heap may hold together where the host gives no heap: a quarter of 2 GiB. */
export const DEFAULT_CACHE_BYTES = 512 * 2 ** 20

export class CacheBudget {
  /** Bytes the caches hold; caches gone with their books count until the next trim. */
  held = 0
  /** Reads so far: an entry's `used` orders it. */
  reads = 0
  /** Reads before the formula computed last: entries read since are not dropped. */
  private since = 0
  private readonly caches = new Set<WeakRef<RangeCache<unknown>>>()
  private sweepAt = 16

  constructor(public cap = DEFAULT_CACHE_BYTES) {}

  join(cache: RangeCache<unknown>): void {
    this.caches.add(new WeakRef(cache))
    // the books computed and dropped leave their references: swept when they double
    if (this.caches.size < this.sweepAt) return
    for (const ref of this.caches) if (!ref.deref()) this.caches.delete(ref)
    this.sweepAt = 2 * this.caches.size + 16
  }

  /** Between formulas: past the cap, the entries least recently read go, down to three quarters of it. */
  fit(): void {
    if (this.held > this.cap) this.trim()
    this.since = this.reads
  }

  private trim(): void {
    const live: RangeCache<unknown>[] = []
    this.held = 0
    for (const ref of this.caches) {
      const c = ref.deref()
      if (!c) this.caches.delete(ref)
      else {
        live.push(c)
        this.held += c.held
      }
    }
    if (this.held <= this.cap) return
    const old: [number, RangeCache<unknown>, string][] = []
    for (const c of live) c.readBefore(this.since, old)
    old.sort((a, b) => a[0] - b[0])
    for (const [, c, key] of old) {
      if (this.held <= this.cap * 0.75) break
      c.delete(key)
    }
  }
}

/** The budget of the books computed in this heap (this thread): a pool's thread sets its cap by `maxHeapMb`. */
export const heapCaches = new CacheBudget()

export class RangeCache<T> extends Map<string, T> {
  private readonly ranges = new Map<string, Entry<T>>()
  private readonly index = new AreaIndex<Entry<T>>()
  /** Bytes this cache holds, by the budget's estimate. */
  held = 0

  constructor(private readonly budget?: CacheBudget) {
    super()
    budget?.join(this as RangeCache<unknown>)
  }

  override get(key: string): T | undefined {
    const entry = this.ranges.get(key)
    if (!entry) return undefined
    if (this.budget) entry.used = ++this.budget.reads
    return entry.value
  }

  /** `size`: bytes the value holds, which a budget counts. */
  override set(key: string, value: T, size = 0): this {
    this.put(key, value, size)
    return super.set(key, value)
  }

  /**
   * Whether a value just made for `key` is worth keeping: it is, when its range was read before since it last changed.
   * The first read leaves only a mark, so ranges read once (a running total's prefixes) fill no cache.
   */
  admits(key: string): boolean {
    if (this.ranges.has(key)) return true
    this.put(key, undefined, 0)
    return false
  }

  private put(key: string, value: T | undefined, size: number): void {
    let entry = this.ranges.get(key)
    if (entry) this.count(-entry.size)
    else {
      const [s, r1, c1, r2, c2] = key.split(':').map(Number)
      entry = { key, s, r1, c1, r2, c2, value, size: 0, used: 0 }
      this.ranges.set(key, entry)
      this.index.add(entry)
    }
    entry.value = value
    entry.size = ENTRY_BYTES + size
    if (this.budget) entry.used = ++this.budget.reads
    this.count(entry.size)
  }

  override delete(key: string): boolean {
    const entry = this.ranges.get(key)
    if (!entry) return false
    this.index.delete(entry)
    this.ranges.delete(key)
    this.count(-entry.size)
    return super.delete(key)
  }

  override clear(): void {
    super.clear()
    this.ranges.clear()
    this.index.clear()
    this.count(-this.held)
  }

  invalidate(area: Area): void {
    this.index.forEach(area, (entry) => this.delete(entry.key))
  }

  /** The entries last read before read `since`, with when, for the budget to drop. */
  readBefore(since: number, out: [number, RangeCache<unknown>, string][]): void {
    for (const entry of this.ranges.values()) if (entry.used <= since) out.push([entry.used, this as RangeCache<unknown>, entry.key])
  }

  private count(bytes: number): void {
    this.held += bytes
    if (this.budget) this.budget.held += bytes
  }
}
