/**
 * Rows on a filtered sheet, where Excel acts on the visible rows (the struct set `hidden`): deleting rows deletes
 * the visible ones among them and inserting rows inserts as many rows as are visible among them, at the first; a
 * copy leaves out the hidden rows of its source (`clip.ts`) and pastes into each run of visible rows of its target
 * as into a target alone. A selection of hidden rows only is taken whole. Every hidden row of a filtered sheet
 * counts, whatever hid it.
 */
import type { BookInput } from '../contract.js'
import { hiddenRows } from '../layout/hidden.js'
import { say } from '../locale/messages.js'
import { cornersText, type Rect } from './area.js'
import { clipped, CLIP } from './clip.js'
import { copyRange } from './copy.js'
import type { Applied, Geometry } from './geometry.js'
import { shift } from './shift.js'
import { sheetOf } from './rewrite.js'
import type { ValuesIn } from './values.js'

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

/** Applies the steps in turn, each to the book the last one gave; the geometry is theirs one after another. */
function chain(book: BookInput, steps: ((b: BookInput) => Applied)[]): Applied {
  let out = book
  const geometry: Geometry = []
  for (const step of steps) {
    const r = step(out)
    if (!r.ok) return r
    out = r.book
    geometry.push(...r.geometry)
  }
  return { ok: true, book: out, geometry }
}

/** `n` rows inserted before `at` (zero-based) on `sheet`: on a filtered sheet as many as are visible among them. */
export function insertRows(book: BookInput, sheet: string, at: number, n: number): Applied {
  const visible = visibleRows(book, sheet, at, at + n - 1)
  return shift(book, sheet, 'r', at, visible?.length || n)
}

/** Rows `at` … `at + n - 1` deleted from `sheet`: on a filtered sheet the visible ones, each run of them from the bottom up. */
export function deleteRows(book: BookInput, sheet: string, at: number, n: number): Applied {
  const visible = visibleRuns(book, sheet, at, at + n - 1)
  if (!visible?.length) return shift(book, sheet, 'r', at, -n)
  return chain(book, visible.reverse().map(([lo, hi]) => (b: BookInput) => shift(b, sheet, 'r', lo, lo - hi - 1)))
}

/**
 * The runs of visible rows a paste into `target` of sheet `to` fills one by one, or null where it fills the target whole:
 * one cell, or no hidden row of a filtered sheet.
 */
export function pasteRuns(book: BookInput, to: string, target: Rect): [number, number][] | null {
  const one = target.r1 === target.r2 && target.c1 === target.c2
  const visible = one ? null : visibleRuns(book, to, target.r1, target.r2)
  return visible?.length ? visible : null
}

/**
 * Copy and paste over filtered sheets. A source over hidden rows of a filtered sheet is copied as its visible rows,
 * packed and as values (`values` gives them). A target of several cells over hidden rows of a filtered sheet takes
 * the copy in each run of its visible rows, as a target alone; a copy that left rows out cannot be pasted over
 * several runs (Excel refuses a selection of several areas pasted onto several areas).
 */
export function copyRows(book: BookInput, from: string, src: Rect, to: string, target: Rect, withMeta: boolean, values: ValuesIn): Applied {
  const clip = clipped(book, from, src, values)
  const [source, at, rect] = clip ? [clip.book, CLIP, clip.src] : [book, from, src]
  const parts = pasteRuns(book, to, target)
  if (!parts) return copyRange(book, at, rect, to, target, withMeta, source)
  if (clip && parts.length > 1)
    return { ok: false, code: 'filtered', message: say(book.locale, { key: 'filteredPaste', sheet: to, area: cornersText(target) }) }
  return chain(book, parts.map(([lo, hi]) => (b: BookInput) => copyRange(b, at, rect, to, { ...target, r1: lo, r2: hi }, withMeta, source)))
}
