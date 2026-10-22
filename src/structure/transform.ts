/**
 * The transforms of references that a copy makes, and the spans of a reference's axis after rows or columns are inserted
 * or deleted: the low layer that the structure operations and the rules' handlers (`cond/structure.ts`,
 * `validation/structure.ts`) share.
 */
import { isFull, LIMIT, span, type Area, type Axis, type Bound } from './area.js'
import type { Outcome, Transform } from './rewrite.js'

/**
 * Relative bounds moved by the copy's offset; a whole column keeps its rows, a whole row its columns. Off the sheet a
 * reference is lost, or with `wrap` comes round its other edge, as a conditional format's formula computes.
 */
export function copyTransform(dr: number, dc: number, wrap = false): Transform {
  return {
    copy: true,
    ref({ sheet, area }): Outcome {
      const next: Area = { ...area }
      for (const [axis, d] of [['r', dr], ['c', dc]] as [Axis, number][]) {
        if (d === 0 || isFull(area[axis], axis)) continue
        const at = (i: number) => (wrap ? (((i + d) % LIMIT[axis]) + LIMIT[axis]) % LIMIT[axis] : i + d)
        const move = (b: Bound): Bound => (b.abs ? b : { i: at(b.i), abs: false })
        const s = span(move(area[axis].lo), move(area[axis].hi))
        if (s.lo.i < 0 || s.hi.i >= LIMIT[axis]) return 'lost'
        next[axis] = s
      }
      return { sheet, area: next }
    },
  }
}

/** One axis of a reference after `n` rows (columns) are inserted before index `at`; null when it leaves the sheet. */
export function insertSpan(lo: number, hi: number, at: number, n: number, max: number, cell: boolean): [number, number] | null {
  if (!cell && lo === 0 && hi === max - 1) return [lo, hi]
  const a = lo >= at ? lo + n : lo
  const b = hi >= at ? hi + n : hi
  if (a > max - 1 || (cell && b > max - 1)) return null
  return [a, Math.min(b, max - 1)]
}

/** One axis of a reference after `n` rows (columns) from index `at` are deleted; null when nothing of it is left. */
export function deleteSpan(lo: number, hi: number, at: number, n: number, max: number, cell: boolean): [number, number] | null {
  const end = at + n
  // an index in the deleted rows goes to `gone`: the first row after them, or the last before them
  const map = (i: number, gone: number) => (i < at ? i : i >= end ? i - n : gone)
  if (lo >= at && hi < end) return null
  if (cell) return [map(lo, lo), map(lo, lo)]
  const a = map(lo, at)
  // a range ending on the sheet's last row keeps ending there
  const b = hi === max - 1 ? hi : map(hi, at - 1)
  return a > b ? null : [a, b]
}
