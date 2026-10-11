/**
 * The cells of the book's input that an array cannot spill over: a value (an empty text too) or a formula.
 * A format or a style alone does not count, as in Excel. Rows are kept sorted by column, so a range of a
 * million rows is checked by binary search.
 */
import type { Rect } from '../structure/area.js'
import { lowerBound } from '../sorted.js'

const rowOf = (row: number) => row

export class Occupied {
  /** sheet → column → rows; a column is sorted before it is searched. */
  private readonly cols = new Map<number, Map<number, number[]>>()
  private readonly unsorted = new Set<number[]>()

  /** Marks a cell; `fresh`: it is not marked yet (building from the input), so the check is skipped. */
  set(s: number, r: number, c: number, on: boolean, fresh = false) {
    if (!fresh && this.has(s, r, c) === on) return
    let sheet = this.cols.get(s)
    if (!sheet) this.cols.set(s, (sheet = new Map()))
    let rows = sheet.get(c)
    if (!rows) sheet.set(c, (rows = []))
    if (on && fresh) {
      rows.push(r)
      this.unsorted.add(rows)
    } else if (on) {
      // an edit puts the row in its place: sorting the column again for each cell made a column of edits quadratic
      rows.splice(lowerBound(this.sorted(rows), r, rowOf), 0, r)
    } else rows.splice(lowerBound(this.sorted(rows), r, rowOf), 1)
  }

  has(s: number, r: number, c: number): boolean {
    const rows = this.cols.get(s)?.get(c)
    return !!rows && this.sorted(rows)[lowerBound(rows, r, rowOf)] === r
  }

  /** The first occupied cell of `x` on sheet `s` (by row, then column) other than (`ar`, `ac`), or null. */
  first(s: number, x: Rect, ar: number, ac: number): [number, number] | null {
    const sheet = this.cols.get(s)
    if (!sheet) return null
    let best: [number, number] | null = null
    for (const [c, rows] of sheet) {
      if (c < x.c1 || c > x.c2) continue
      this.sorted(rows)
      for (let i = lowerBound(rows, x.r1, rowOf); i < rows.length && rows[i] <= x.r2; i++) {
        if (rows[i] === ar && c === ac) continue
        if (!best || rows[i] < best[0] || (rows[i] === best[0] && c < best[1])) best = [rows[i], c]
        break
      }
    }
    return best
  }

  private sorted(rows: number[]): number[] {
    if (this.unsorted.delete(rows)) rows.sort((a, b) => a - b)
    return rows
  }
}
