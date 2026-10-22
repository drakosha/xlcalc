/** References as rectangles with their `$` flags, and their text in the form Excel writes it. */
import { colToIndex, formatAddress, indexToCol, MAX_COLS, MAX_ROWS, parseAddress } from '../model.js'
import type { RangeRef as SyntaxRef } from '../syntax.js'

/** One bound: a zero-based index and whether it is absolute (`$`). */
export interface Bound {
  i: number
  abs: boolean
}

export interface Span {
  lo: Bound
  hi: Bound
}

/** A reference's rectangle; a whole column has its rows spanning the sheet (and absolute), a whole row its columns. */
export interface Area {
  kind: SyntaxRef['kind']
  r: Span
  c: Span
}

export type Axis = 'r' | 'c'

export const LIMIT: Record<Axis, number> = { r: MAX_ROWS, c: MAX_COLS }

const full = (axis: Axis): Span => ({ lo: { i: 0, abs: true }, hi: { i: LIMIT[axis] - 1, abs: true } })

/** Both bounds in order; Excel orders each axis of `B5:A1` on its own, carrying the `$` with the bound. */
export const span = (a: Bound, b: Bound): Span => (a.i <= b.i ? { lo: a, hi: b } : { lo: b, hi: a })

const colBound = (abs: string, letters: string): Bound => ({ i: colToIndex(letters), abs: abs === '$' })
const rowBound = (abs: string, digits: string): Bound => ({ i: Number(digits) - 1, abs: abs === '$' })

const CELL = /^(\$?)([A-Z]+)(\$?)(\d+)$/
const COLS = /^(\$?)([A-Z]+):(\$?)([A-Z]+)$/
const ROWS = /^(\$?)(\d+):(\$?)(\d+)$/

/** The rectangle of a parsed reference (its text carries the `$` flags). */
export function decode(ref: SyntaxRef): Area {
  const t = ref.text
  if (ref.kind === 'cols') {
    const m = COLS.exec(t)!
    return { kind: 'cols', r: full('r'), c: span(colBound(m[1], m[2]), colBound(m[3], m[4])) }
  }
  if (ref.kind === 'rows') {
    const m = ROWS.exec(t)!
    return { kind: 'rows', r: span(rowBound(m[1], m[2]), rowBound(m[3], m[4])), c: full('c') }
  }
  const [a, b = a] = t.split(':').map((x) => CELL.exec(x)!)
  return { kind: ref.kind, r: span(rowBound(a[3], a[4]), rowBound(b[3], b[4])), c: span(colBound(a[1], a[2]), colBound(b[1], b[2])) }
}

const colText = (b: Bound) => (b.abs ? '$' : '') + indexToCol(b.i)
const rowText = (b: Bound) => (b.abs ? '$' : '') + String(b.i + 1)

export const isFull = (s: Span, axis: Axis) => s.lo.i === 0 && s.hi.i === LIMIT[axis] - 1

/**
 * The reference text as Excel writes it: a range over every row is a column range (`A:A`), over every
 * column a row range (`1:1`), over the whole sheet the rows `1:1048576`.
 */
export function encode(a: Area): string {
  if (a.kind === 'cell') return colText(a.c.lo) + rowText(a.r.lo)
  if (isFull(a.c, 'c')) return `${rowText(a.r.lo)}:${rowText(a.r.hi)}`
  if (isFull(a.r, 'r')) return `${colText(a.c.lo)}:${colText(a.c.hi)}`
  return `${colText(a.c.lo)}${rowText(a.r.lo)}:${colText(a.c.hi)}${rowText(a.r.hi)}`
}

export const sameSpan = (x: Span, y: Span) => x.lo.i === y.lo.i && x.hi.i === y.hi.i && x.lo.abs === y.lo.abs && x.hi.abs === y.hi.abs

export const sameArea = (x: Area, y: Area) => sameSpan(x.r, y.r) && sameSpan(x.c, y.c)

/** The area with one axis replaced; bounds keep their `$` flags. */
export function withSpan(a: Area, axis: Axis, lo: number, hi: number): Area {
  const s = a[axis]
  const next: Span = { lo: { i: lo, abs: s.lo.abs }, hi: { i: hi, abs: s.hi.abs } }
  return axis === 'r' ? { ...a, r: next } : { ...a, c: next }
}

/** A plain rectangle, zero-based and inclusive. */
export interface Rect {
  r1: number
  c1: number
  r2: number
  c2: number
}

export const rectOf = (a: Area): Rect => ({ r1: a.r.lo.i, r2: a.r.hi.i, c1: a.c.lo.i, c2: a.c.hi.i })

export const lo = (x: Rect, axis: Axis) => (axis === 'r' ? x.r1 : x.c1)
export const hi = (x: Rect, axis: Axis) => (axis === 'r' ? x.r2 : x.c2)

/** The rectangle with its bounds along `axis` replaced. */
export const withBounds = (x: Rect, axis: Axis, a: number, b: number): Rect => (axis === 'r' ? { ...x, r1: a, r2: b } : { ...x, c1: a, c2: b })

export const contains = (outer: Rect, inner: Rect) => inner.r1 >= outer.r1 && inner.r2 <= outer.r2 && inner.c1 >= outer.c1 && inner.c2 <= outer.c2

export const meets = (x: Rect, y: Rect) => x.r1 <= y.r2 && y.r1 <= x.r2 && x.c1 <= y.c2 && y.c1 <= x.c2

/** The range grown by the merged areas it meets, each taken in whole, until it meets no other. */
export function grown(x: Rect, merges: Rect[]): Rect {
  let out = x
  for (let more = true; more; ) {
    more = false
    for (const m of merges)
      if (meets(m, out) && (m.r1 < out.r1 || m.c1 < out.c1 || m.r2 > out.r2 || m.c2 > out.c2)) {
        out = { r1: Math.min(out.r1, m.r1), c1: Math.min(out.c1, m.c1), r2: Math.max(out.r2, m.r2), c2: Math.max(out.c2, m.c2) }
        more = true
      }
  }
  return out
}

/** Whether the cell at row `r`, column `c` lies in the rectangle. */
export const hasCell = (x: Rect, r: number, c: number) => r >= x.r1 && r <= x.r2 && c >= x.c1 && c <= x.c2

/** The number of a rectangle's cells. */
export const cellCount = (x: Rect) => (x.r2 - x.r1 + 1) * (x.c2 - x.c1 + 1)

/** A rectangle's cells, row by row. */
export function* rectCells(x: Rect): Generator<string> {
  for (let r = x.r1; r <= x.r2; r++) for (let c = x.c1; c <= x.c2; c++) yield formatAddress(r, c)
}

/**
 * A rectangle's cells among a sheet's (`cells`, its addresses): listed row by row up to `max` cells, those of the sheet
 * within it past that, which is then cheaper. The listed ones need not be the sheet's.
 */
export function* cellsWithin(x: Rect, cells: () => Iterable<string>, max = 4096): Generator<string> {
  if (cellCount(x) <= max) {
    yield* rectCells(x)
    return
  }
  for (const a of cells()) {
    const { row, col } = parseAddress(a)
    if (hasCell(x, row, col)) yield a
  }
}

/** The cells two rectangles share, or null when they do not meet. */
export function intersect(x: Rect, y: Rect): Rect | null {
  const z = { r1: Math.max(x.r1, y.r1), r2: Math.min(x.r2, y.r2), c1: Math.max(x.c1, y.c1), c2: Math.min(x.c2, y.c2) }
  return z.r1 <= z.r2 && z.c1 <= z.c2 ? z : null
}

export const shiftRect = (x: Rect, dr: number, dc: number): Rect => ({ r1: x.r1 + dr, r2: x.r2 + dr, c1: x.c1 + dc, c2: x.c2 + dc })

/** `x` without the cells of `y`: up to four rectangles, above, below, left and right of it. */
export function minus(x: Rect, y: Rect): Rect[] {
  if (!meets(x, y)) return [x]
  const out: Rect[] = []
  if (x.r1 < y.r1) out.push({ ...x, r2: y.r1 - 1 })
  if (x.r2 > y.r2) out.push({ ...x, r1: y.r2 + 1 })
  const [r1, r2] = [Math.max(x.r1, y.r1), Math.min(x.r2, y.r2)]
  if (x.c1 < y.c1) out.push({ r1, r2, c1: x.c1, c2: y.c1 - 1 })
  if (x.c2 > y.c2) out.push({ r1, r2, c1: y.c2 + 1, c2: x.c2 })
  return out
}

/** Rectangles without the cells of `cut`. */
export const subtract = (rects: Rect[], cut: Rect): Rect[] => rects.flatMap((x) => minus(x, cut))

/** The top-left cell of the rectangle around `xs`, which need not be a cell of theirs: the cell a rule's formulas are written for. */
export const cornerOf = (xs: Rect[]): { r: number; c: number } => ({ r: Math.min(...xs.map((x) => x.r1)), c: Math.min(...xs.map((x) => x.c1)) })

/** A rectangle's text from corner to corner, `B2:D9`, one cell too (`B2:B2`): a merged area, a table, a spill. */
export const cornersText = (x: Rect) => `${formatAddress(x.r1, x.c1)}:${formatAddress(x.r2, x.c2)}`

/** A rectangle's text: `B2:D9`, or `B2` for one cell. */
export const rectText = (x: Rect) => (x.r1 === x.r2 && x.c1 === x.c2 ? formatAddress(x.r1, x.c1) : cornersText(x))

/** A rectangle as a part of a range: a cell, a range, whole columns (`C:D`) or whole rows (`2:5`), as Excel writes references. */
export function areaText(x: Rect): string {
  if (x.r1 === 0 && x.r2 === MAX_ROWS - 1) return `${indexToCol(x.c1)}:${indexToCol(x.c2)}`
  if (x.c1 === 0 && x.c2 === MAX_COLS - 1) return `${x.r1 + 1}:${x.r2 + 1}`
  return rectText(x)
}

const RANGE_ARG = /^(?:([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?|([A-Z]{1,3}):([A-Z]{1,3})|(\d+):(\d+))$/

/** An operation's range argument (`B2`, `B2:D9`, `B:D`, `2:9`), or null when it is not one or leaves the sheet. */
export function parseRange(text: string): Rect | null {
  const m = RANGE_ARG.exec(text)
  if (!m) return null
  let x: Rect
  if (m[1]) {
    const [c1, r1] = [colToIndex(m[1]), Number(m[2]) - 1]
    const [c2, r2] = m[3] ? [colToIndex(m[3]), Number(m[4]) - 1] : [c1, r1]
    x = { r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2) }
  } else if (m[5]) {
    const [a, b] = [colToIndex(m[5]), colToIndex(m[6])]
    x = { r1: 0, r2: MAX_ROWS - 1, c1: Math.min(a, b), c2: Math.max(a, b) }
  } else {
    const [a, b] = [Number(m[7]) - 1, Number(m[8]) - 1]
    x = { r1: Math.min(a, b), r2: Math.max(a, b), c1: 0, c2: MAX_COLS - 1 }
  }
  return x.r1 >= 0 && x.c1 >= 0 && x.r2 < MAX_ROWS && x.c2 < MAX_COLS ? x : null
}
