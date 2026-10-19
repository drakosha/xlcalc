/** The rows a filter leaves visible on a sheet, which the structure operations on a filtered sheet act on (`filtered.ts`, `clip.ts`). */
import type { BookInput } from '../contract.js'
import { hiddenRows } from '../layout/hidden.js'
import { sheetOf } from './rewrite.js'

/**
 * The runs of rows of [lo, hi] a filter leaves visible on `sheet`, ascending, read from its spans of hidden rows; null
 * when the sheet is not filtered or none of them is hidden.
 */
export function visibleRuns(book: BookInput, sheet: string, lo: number, hi: number): [number, number][] | null {
  const s = sheetOf(book, sheet)
  const h = s.filtered ? hiddenRows(s) : undefined
  if (!h?.spans.some(([a, b]) => a <= hi && b >= lo)) return null
  const out: [number, number][] = []
  let r = lo
  for (const [a, b] of h.spans) {
    if (a > hi) break
    if (a > r) out.push([r, a - 1])
    r = Math.max(r, b + 1)
  }
  if (r <= hi) out.push([r, hi])
  return out
}

/** The rows of [lo, hi] a filter leaves visible on `sheet`, or null when the sheet is not filtered or none of them is hidden. */
export function visibleRows(book: BookInput, sheet: string, lo: number, hi: number): number[] | null {
  return visibleRuns(book, sheet, lo, hi)?.flatMap(([a, b]) => Array.from({ length: b - a + 1 }, (_, k) => a + k)) ?? null
}
