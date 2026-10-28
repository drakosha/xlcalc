/**
 * The cells of a sheet whose formula calls SUBTOTAL or AGGREGATE (row · MAX_COLS + column), which those two skip, with
 * a count by column: whether a range can hold one is told by its columns, not by a pass over every mark of the sheet.
 */
import { MAX_COLS } from './model.js'

export class NestedMarks extends Set<number> {
  private readonly columns = new Map<number, number>()

  add(pos: number): this {
    if (!this.has(pos)) {
      const c = pos % MAX_COLS
      this.columns.set(c, (this.columns.get(c) ?? 0) + 1)
    }
    return super.add(pos)
  }

  delete(pos: number): boolean {
    if (!super.delete(pos)) return false
    const c = pos % MAX_COLS
    const n = this.columns.get(c)! - 1
    if (n) this.columns.set(c, n)
    else this.columns.delete(c)
    return true
  }

  /** Whether a mark lies in a column from `c1` to `c2`. */
  inColumns(c1: number, c2: number): boolean {
    if (c2 - c1 + 1 > this.columns.size) {
      for (const c of this.columns.keys()) if (c >= c1 && c <= c2) return true
      return false
    }
    for (let c = c1; c <= c2; c++) if (this.columns.has(c)) return true
    return false
  }
}
