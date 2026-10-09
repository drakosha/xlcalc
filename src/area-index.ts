/** Sheet/column index shared by dependency listeners and range caches. */
import type { Area } from './spill/state.js'
import { meets } from './structure/area.js'

const WIDE = 64

export class AreaIndex<T extends Area> {
  private readonly columns = new Map<number, Map<number, Set<T>>>()
  private readonly wide = new Map<number, Set<T>>()

  add(entry: T): void {
    const { s, c1, c2 } = entry
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

  delete(entry: T): void {
    const { s, c1, c2 } = entry
    if (c2 - c1 >= WIDE) {
      const set = this.wide.get(s)
      if (!set) return
      set.delete(entry)
      if (!set.size) this.wide.delete(s)
    } else {
      const cols = this.columns.get(s)
      if (!cols) return
      for (let c = c1; c <= c2; c++) {
        const set = cols.get(c)
        if (!set) continue
        set.delete(entry)
        if (!set.size) cols.delete(c)
      }
      if (!cols.size) this.columns.delete(s)
    }
  }

  clear(): void {
    this.columns.clear()
    this.wide.clear()
  }

  /** A hit may occur in several columns. Visitors can deduplicate it or delete it from this index. */
  forEach(area: Area, visit: (entry: T) => void): void {
    const scan = (entries: Set<T> | undefined) => {
      if (entries) for (const entry of entries) if (meets(entry, area)) visit(entry)
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
