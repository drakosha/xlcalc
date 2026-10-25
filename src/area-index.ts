/** Sheet/column index shared by dependency listeners and range caches. */
import type { Area } from './spill/state.js'
import { meets } from './structure/area.js'

const WIDE = 64
/** Rows of a bucket of a column: 2^SHIFT. */
const SHIFT = 6

/**
 * A column's entries in the order they were added, how many of them are taller than 2^SHIFT rows, and the others by the
 * buckets of rows they reach: a lookup of one bucket's rows in a column without tall entries reads that bucket alone,
 * which lists its entries in the same order.
 */
interface Column<T> {
  all: Set<T>
  tall: number
  rows: Map<number, Set<T>>
}

export class AreaIndex<T extends Area> {
  private readonly columns = new Map<number, Map<number, Column<T>>>()
  private readonly wide = new Map<number, Set<T>>()

  add(entry: T): void {
    const { s, r1, r2, c1, c2 } = entry
    if (c2 - c1 >= WIDE) {
      let set = this.wide.get(s)
      if (!set) this.wide.set(s, set = new Set())
      set.add(entry)
    } else {
      let cols = this.columns.get(s)
      if (!cols) this.columns.set(s, cols = new Map())
      for (let c = c1; c <= c2; c++) {
        let col = cols.get(c)
        if (!col) cols.set(c, col = { all: new Set(), tall: 0, rows: new Map() })
        const size = col.all.size
        if (col.all.add(entry).size === size) continue
        if (r2 - r1 >> SHIFT) col.tall++
        else for (let q = r1 >> SHIFT; q <= r2 >> SHIFT; q++) {
          let set = col.rows.get(q)
          if (!set) col.rows.set(q, set = new Set())
          set.add(entry)
        }
      }
    }
  }

  delete(entry: T): void {
    const { s, r1, r2, c1, c2 } = entry
    if (c2 - c1 >= WIDE) {
      const set = this.wide.get(s)
      if (!set) return
      set.delete(entry)
      if (!set.size) this.wide.delete(s)
    } else {
      const cols = this.columns.get(s)
      if (!cols) return
      for (let c = c1; c <= c2; c++) {
        const col = cols.get(c)
        if (!col?.all.delete(entry)) continue
        if (r2 - r1 >> SHIFT) col.tall--
        else for (let q = r1 >> SHIFT; q <= r2 >> SHIFT; q++) {
          const set = col.rows.get(q)!
          set.delete(entry)
          if (!set.size) col.rows.delete(q)
        }
        if (!col.all.size) cols.delete(c)
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
    const scan = (col: Column<T> | undefined) => {
      const q = area.r1 >> SHIFT
      const entries = col && (col.tall || q !== area.r2 >> SHIFT ? col.all : col.rows.get(q))
      if (entries) for (const entry of entries) if (meets(entry, area)) visit(entry)
    }
    const cols = this.columns.get(area.s)
    if (cols) {
      if (area.c2 - area.c1 + 1 > cols.size) {
        for (const [c, col] of cols) if (c >= area.c1 && c <= area.c2) scan(col)
      } else for (let c = area.c1; c <= area.c2; c++) scan(cols.get(c))
    }
    const wide = this.wide.get(area.s)
    if (wide) for (const entry of wide) if (meets(entry, area)) visit(entry)
  }
}
