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
  /** A mark (`admits`) is distinct from a stored undefined. */
  value: T | typeof MARK
  /** Bytes by the budget's estimate, the entry's own with its value's. */
  size: number
  /** The budget's read count when it was last read. */
  used: number
}

const MARK = Symbol('range first read')

/** Bytes of an entry besides its value, a mark's too: its key, its bounds and its places in the maps, as measured on Node. */
const ENTRY_BYTES = 256

/** Retained storage estimates shared by materialized range/index caches and optional-index selection. */
export const VALUE_BYTES = 8
export const POSITION_BYTES = 16
export const KEY_BYTES = 80

/** Bytes the range caches of the books in one heap may hold together where the host gives no heap: a quarter of 2 GiB. */
const DEFAULT_CACHE_BYTES = 512 * 2 ** 20

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

/** How the cache that keeps an array counts what is kept with it (`RangeCache.carries`), by the array, weakly. */
const carried = new WeakMap<object, (bytes: number) => void>()

/** Bytes kept with an array (a lookup index over it), counted in the budget of the cache that keeps the array, if one does. */
export function keptWith(value: object, bytes: number): void {
  carried.get(value)?.(bytes)
}

export class RangeCache<T> implements Map<string, T> {
  readonly [Symbol.toStringTag] = 'RangeCache'
  private readonly ranges = new Map<string, Entry<T>>()
  private readonly index = new AreaIndex<Entry<T>>()
  /** Bytes this cache holds, by the budget's estimate. */
  held = 0
  private stored = 0

  get size(): number { return this.stored }
  has(key: string): boolean {
    const entry = this.ranges.get(key)
    return entry !== undefined && entry.value !== MARK
  }

  *entries(): MapIterator<[string, T]> {
    for (const entry of this.ranges.values()) if (entry.value !== MARK) yield [entry.key, entry.value]
  }
  *keys(): MapIterator<string> {
    for (const [key] of this.entries()) yield key
  }
  *values(): MapIterator<T> {
    for (const [, value] of this.entries()) yield value
  }
  [Symbol.iterator](): MapIterator<[string, T]> { return this.entries() }
  forEach(visit: (value: T, key: string, map: Map<string, T>) => void, thisArg?: unknown): void {
    for (const [key, value] of this.entries()) visit.call(thisArg, value, key, this)
  }

  constructor(private readonly budget?: CacheBudget) {
    budget?.join(this as RangeCache<unknown>)
  }

  get(key: string): T | undefined {
    const entry = this.ranges.get(key)
    if (!entry) return undefined
    if (this.budget) entry.used = ++this.budget.reads
    return entry.value === MARK ? undefined : entry.value
  }

  /** `size`: bytes the value holds, which a budget counts. */
  set(key: string, value: T, size = 0): this {
    this.put(key, value, size)
    return this
  }

  /**
   * Whether a value just made for `key` is worth keeping: it is, when its range was read before since it last changed.
   * The first read leaves only a mark, so ranges read once (a running total's prefixes) fill no cache.
   */
  admits(key: string): boolean {
    if (this.ranges.has(key)) return true
    this.put(key, MARK, 0)
    return false
  }

  private put(key: string, value: T | typeof MARK, size: number): void {
    let entry = this.ranges.get(key)
    if (entry) {
      this.count(-entry.size)
      if (entry.value === MARK && value !== MARK) {
        // A promoted mark joins stored values at the end, as Map.set did in the former value map.
        this.ranges.delete(key)
        this.ranges.set(key, entry)
      }
    } else {
      const [s, r1, c1, r2, c2] = key.split(':').map(Number)
      entry = { key, s, r1, c1, r2, c2, value: MARK, size: 0, used: 0 }
      this.ranges.set(key, entry)
      this.index.add(entry)
    }
    if (value !== MARK && entry.value === MARK) this.stored++
    entry.value = value
    entry.size = ENTRY_BYTES + size
    if (this.budget) entry.used = ++this.budget.reads
    this.count(entry.size)
  }

  /** Lets `value`, kept under `key`, count what is kept with it (`keptWith`) while it is kept there; it goes with it. */
  carries(key: string, value: T & object): void {
    carried.set(value, (bytes) => {
      const entry = this.ranges.get(key)
      if (entry?.value !== value) return
      entry.size += bytes
      this.count(bytes)
    })
  }

  delete(key: string): boolean {
    const entry = this.ranges.get(key)
    if (!entry) return false
    this.index.delete(entry)
    this.ranges.delete(key)
    this.count(-entry.size)
    if (entry.value === MARK) return false
    this.stored--
    return true
  }

  clear(): void {
    this.stored = 0
    this.ranges.clear()
    this.index.clear()
    this.count(-this.held)
  }

  invalidate(area: Area): void {
    // every write of a value drops what meets its cell: an empty cache, the usual one, has nothing to look up
    if (!this.ranges.size) return
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
