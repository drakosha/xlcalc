/** References as rectangles with their `$` flags, and their text in the form Excel writes it. */
import { colToIndex, indexToCol, MAX_COLS, MAX_ROWS } from '../model.js'
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

export const contains = (outer: Rect, inner: Rect) => inner.r1 >= outer.r1 && inner.r2 <= outer.r2 && inner.c1 >= outer.c1 && inner.c2 <= outer.c2

export const meets = (x: Rect, y: Rect) => x.r1 <= y.r2 && y.r1 <= x.r2 && x.c1 <= y.c2 && y.c1 <= x.c2

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
