/**
 * Functions that reshape an array (Excel 365): TRANSPOSE, CHOOSECOLS, CHOOSEROWS, TAKE, DROP, EXPAND, HSTACK,
 * VSTACK, TOCOL, TOROW, WRAPROWS, WRAPCOLS. This engine's own. An empty answer is #CALC! in Excel, #VALUE!
 * here (`EMPTY`).
 */
import type { Arg, Cx, FnDef, Grid, Segment } from '../../args.js'
import { isMany } from '../../arrays/lift.js'
import { Arr, EMPTY, fromGrid } from '../../arrays/value.js'
import { E, isErr, Rng, toNum, type CErr, type V } from '../../values.js'

/** Rows `rs` and columns `cs` (zero-based) of an array. */
function pick(a: Arr, rs: readonly number[], cs: readonly number[]): Arr {
  const v: V[] = []
  for (const r of rs) for (const c of cs) v.push(a.at(r, c))
  return new Arr(rs.length, cs.length, rs.length, cs.length, v)
}

const range = (from: number, to: number) => Array.from({ length: Math.max(0, to - from) }, (_, i) => from + i)

/** CHOOSECOLS/CHOOSEROWS indexes: numbers (an array gives each of its values), negative from the end, 0 or past the end #VALUE!. */
function indexes(segs: readonly Segment[], n: number): number[] | CErr {
  const out: number[] = []
  for (const s of segs)
    for (const v of Array.isArray(s) ? (s as readonly V[]) : [s as V]) {
      const k = toNum(v)
      if (isErr(k)) return k
      const i = Math.trunc(k)
      if (i === 0 || Math.abs(i) > n) return E.VALUE
      out.push(i > 0 ? i - 1 : n + i)
    }
  return out
}

/** TAKE and DROP: the first (last, when negative) `n` of `size`, as zero-based indexes. */
function part(size: number, n: number | undefined, take: boolean): number[] {
  if (n === undefined) return range(0, size)
  const k = Math.min(Math.abs(Math.trunc(n)), size)
  if (take) return n >= 0 ? range(0, k) : range(size - k, size)
  return n >= 0 ? range(k, size) : range(0, size - k)
}

/** TAKE and DROP: of a reference a reference (COUNTA skips its blank cells), of an array an array. */
function takeDrop(take: boolean): FnDef {
  return {
    args: 'R n n?',
    cx: true,
    array: true,
    fn: (cx: Cx, x: Arg, rows: number, cols?: number) => {
      if (isErr(x)) return x
      if (take && (Math.trunc(rows) === 0 || (cols !== undefined && Math.trunc(cols) === 0))) return EMPTY
      const g = cx.grid(x)
      const rs = part(g.rows, rows, take)
      const cs = part(g.cols, cols, take)
      if (!rs.length || !cs.length) return EMPTY
      if (x instanceof Rng) return new Rng(x.s, x.r1 + rs[0], x.c1 + cs[0], x.r1 + rs[rs.length - 1], x.c1 + cs[cs.length - 1])
      return pick(fromGrid(g), rs, cs)
    },
  }
}

/** An argument of HSTACK, VSTACK: a value, a reference or an array, as an array (an error is a 1 × 1 one). */
const stackPart = (cx: Cx, x: Arg): Arr => (x instanceof Arr ? x : x instanceof Rng ? fromGrid(cx.grid(x)) : new Arr(1, 1, 1, 1, [x]))

/** Arrays side by side (`across`) or one under another; shorter ones are padded with #N/A. */
function stack(across: boolean): FnDef {
  return {
    args: 'p',
    rep: 1,
    cx: true,
    array: true,
    fn: (cx: Cx, xs: Arg[]) => {
      const parts = xs.map((x) => stackPart(cx, x))
      const rows = across ? Math.max(...parts.map((p) => p.rows)) : parts.reduce((s, p) => s + p.rows, 0)
      const cols = across ? parts.reduce((s, p) => s + p.cols, 0) : Math.max(...parts.map((p) => p.cols))
      const v: V[] = new Array(rows * cols).fill(E.NA)
      let at = 0
      for (const p of parts) {
        for (let r = 0; r < p.rows; r++) for (let c = 0; c < p.cols; c++) v[(across ? r : at + r) * cols + (across ? at + c : c)] = p.at(r, c)
        at += across ? p.cols : p.rows
      }
      return new Arr(rows, cols, rows, cols, v)
    },
  }
}

/** TOCOL, TOROW: the values row by row (column by column), without blanks (ignore 1), errors (2) or both (3). */
function flatten(column: boolean): FnDef {
  return {
    args: 'G i? b?',
    array: true,
    fn: (g: Grid, ignore = 0, byCol = false) => {
      if (ignore < 0 || ignore > 3) return E.VALUE
      const a = fromGrid(g)
      const v: V[] = []
      const [outer, inner] = byCol ? [a.cols, a.rows] : [a.rows, a.cols]
      // a whole column without blanks: only its occupied part
      const stop = ignore & 1 && a.fill === null ? (byCol ? Math.min(outer, a.w) : Math.min(outer, a.h)) : outer
      for (let i = 0; i < stop; i++)
        for (let j = 0; j < inner; j++) {
          const x = byCol ? a.at(j, i) : a.at(i, j)
          if ((x === null && ignore & 1) || (isErr(x) && ignore & 2)) continue
          v.push(x)
        }
      if (!v.length) return EMPTY
      return column ? new Arr(v.length, 1, v.length, 1, v) : new Arr(1, v.length, 1, v.length, v)
    },
  }
}

/** WRAPROWS, WRAPCOLS: a row or column cut into rows (columns) of `n`, the last padded (#N/A by default). */
function wrap(rows: boolean): FnDef {
  return {
    args: 'G n a?',
    array: true,
    fn: (g: Grid, count: number, pad: V = E.NA) => {
      if (g.rows > 1 && g.cols > 1) return E.VALUE
      const n = Math.trunc(count)
      if (n < 1) return E.NUM
      const a = fromGrid(g)
      const len = a.rows * a.cols
      const lines = Math.ceil(len / n)
      const v: V[] = []
      for (let i = 0; i < lines; i++) for (let j = 0; j < n; j++) v.push(i * n + j < len ? a.at(a.rows === 1 ? 0 : i * n + j, a.rows === 1 ? i * n + j : 0) : pad)
      const out = new Arr(lines, n, lines, n, v)
      return rows ? out : transpose(out)
    },
  }
}

function transpose(a: Arr): Arr {
  const v: V[] = new Array(a.h * a.w)
  for (let r = 0; r < a.h; r++) for (let c = 0; c < a.w; c++) v[c * a.h + r] = a.v[r * a.w + c]
  return new Arr(a.cols, a.rows, a.w, a.h, v, a.fill)
}

export const functions: Record<string, FnDef> = {
  CHOOSECOLS: { args: 'G V', rep: 1, array: true, fn: (g: Grid, segs: Segment[]) => {
    const cs = indexes(segs, g.cols)
    return isErr(cs) ? cs : pick(fromGrid(g), range(0, g.rows), cs)
  } },
  CHOOSEROWS: { args: 'G V', rep: 1, array: true, fn: (g: Grid, segs: Segment[]) => {
    const rs = indexes(segs, g.rows)
    return isErr(rs) ? rs : pick(fromGrid(g), rs, range(0, g.cols))
  } },
  DROP: takeDrop(false),
  EXPAND: {
    args: 'G n n? a?',
    array: true,
    fn: (g: Grid, rows: number, cols?: number, pad: V = E.NA) => {
      const r = Math.trunc(rows)
      const c = cols === undefined ? g.cols : Math.trunc(cols)
      if (r < g.rows || c < g.cols) return E.VALUE
      const a = fromGrid(g)
      const v: V[] = []
      for (let i = 0; i < r; i++) for (let j = 0; j < c; j++) v.push(i < a.rows && j < a.cols ? a.at(i, j) : pad)
      return new Arr(r, c, r, c, v)
    },
  },
  HSTACK: stack(true),
  TAKE: takeDrop(true),
  TOCOL: flatten(true),
  TOROW: flatten(false),
  // a value or a one-cell reference is that value (TYPE(TRANSPOSE(5)) is 1), an array of one value stays an array
  TRANSPOSE: { args: 'R', cx: true, array: true, fn: (cx: Cx, x: Arg) => (x instanceof Arr || isMany(x) ? transpose(fromGrid(cx.grid(x))) : cx.scalar(x)) },
  VSTACK: stack(false),
  WRAPCOLS: wrap(false),
  WRAPROWS: wrap(true),
}
