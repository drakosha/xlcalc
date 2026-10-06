/** Merged areas of a sheet: `B2:D3`, top-left to bottom-right, two cells or more, none overlapping another. */
import { colToIndex, formatAddress, MAX_COLS, MAX_ROWS } from '../model.js'
import { meets, type Rect } from '../structure/area.js'

const MERGE = /^([A-Z]{1,3})([1-9][0-9]{0,6}):([A-Z]{1,3})([1-9][0-9]{0,6})$/

/** The area of a merge, or null when it is not one: malformed, reversed, one cell, past the sheet. */
export function parseMerge(text: string): Rect | null {
  const m = MERGE.exec(text)
  if (!m) return null
  const x: Rect = { r1: Number(m[2]) - 1, c1: colToIndex(m[1]), r2: Number(m[4]) - 1, c2: colToIndex(m[3]) }
  const ok = x.r1 <= x.r2 && x.c1 <= x.c2 && (x.r1 < x.r2 || x.c1 < x.c2) && x.r2 < MAX_ROWS && x.c2 < MAX_COLS
  return ok ? x : null
}

export const mergeText = (x: Rect) => `${formatAddress(x.r1, x.c1)}:${formatAddress(x.r2, x.c2)}`

/** The first two merges that overlap, by index, or null (sorted by top row, each meets only those above it that reach down). */
export function overlapping(xs: Rect[]): [number, number] | null {
  const order = xs.map((_, i) => i).sort((a, b) => xs[a].r1 - xs[b].r1)
  const open: number[] = []
  for (const i of order) {
    for (let k = open.length - 1; k >= 0; k--) {
      const j = open[k]
      if (xs[j].r2 < xs[i].r1) open.splice(k, 1)
      else if (meets(xs[i], xs[j])) return [Math.min(i, j), Math.max(i, j)]
    }
    open.push(i)
  }
  return null
}
