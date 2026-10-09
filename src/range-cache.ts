/** Cached ranges with their bounds and a sheet/column index, so a write visits only ranges in its columns. */
import type { Area } from './spill/state.js'
import { meets } from './structure/area.js'

interface Entry extends Area { key: string }
const WIDE = 64

export class RangeCache<T> extends Map<string, T> {
  private readonly ranges = new Map<string, Entry>()
  private readonly columns = new Map<number, Map<number, Set<Entry>>>()
  private readonly wide = new Map<number, Set<Entry>>()

  constructor() { super() }

  override set(key: string, value: T): this {
    if (!this.has(key)) {
      const [s, r1, c1, r2, c2] = key.split(':').map(Number)
      const entry = { key, s, r1, c1, r2, c2 }
      this.ranges.set(key, entry)
      if (c2 - c1 >= WIDE) {
        let set = this.wide.get(s)
        if (!set) this.wide.set(s, set = new Set())
        set.add(entry)
      } else {
        let cols = this.columns.get(s)
        if (!cols) this.columns.set(s, cols = new Map())
        for (let c = c1; c <= c2; c++) {
          let set = cols.get(c)
          if (!set) cols.set(c, set = new Set())
          set.add(entry)
        }
      }
    }
    return super.set(key, value)
  }

  override delete(key: string): boolean {
    const entry = this.ranges.get(key)
    if (!entry) return false
    const { s, c1, c2 } = entry
    if (c2 - c1 >= WIDE) {
      const set = this.wide.get(s)!
      set.delete(entry)
      if (!set.size) this.wide.delete(s)
    } else {
      const cols = this.columns.get(s)!
      for (let c = c1; c <= c2; c++) {
        const set = cols.get(c)!
        set.delete(entry)
        if (!set.size) cols.delete(c)
      }
      if (!cols.size) this.columns.delete(s)
    }
    this.ranges.delete(key)
    return super.delete(key)
  }

  override clear(): void {
    super.clear()
    this.ranges.clear()
    this.columns.clear()
    this.wide.clear()
  }

  invalidate(area: Area): void {
    const scan = (entries: Set<Entry> | undefined) => {
      if (entries) for (const entry of entries) if (meets(entry, area)) this.delete(entry.key)
    }
    const cols = this.columns.get(area.s)
    if (cols) {
      if (area.c2 - area.c1 + 1 > cols.size) {
        for (const [c, entries] of cols) if (c >= area.c1 && c <= area.c2) scan(entries)
      } else for (let c = area.c1; c <= area.c2; c++) scan(cols.get(c))
    }
    scan(this.wide.get(area.s))
  }
}
