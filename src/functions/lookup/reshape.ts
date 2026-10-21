/**
 * Functions that reshape an array (Excel 365): TRANSPOSE, CHOOSECOLS, CHOOSEROWS, TAKE, DROP, EXPAND, HSTACK,
 * VSTACK, TOCOL, TOROW, WRAPROWS, WRAPCOLS. This engine's own. An empty answer is #CALC! in Excel, #VALUE!
 * here (`EMPTY`).
 */
import type { Arg, Cx, FnDef, Grid, Segment } from '../../args.js'
import { isMany } from '../../arrays/lift.js'
import { Arr, asArray, EMPTY, fromGrid, selectArray } from '../../arrays/value.js'
import { allot, spend } from '../../exec/budget.js'
import type { Rules } from '../../locale/rules.js'
import { Areas, E, isErr, Rng, type CErr, type V } from '../../values.js'
import { NEWER_MAX_ARGS } from '../common.js'

/** Rows `rs` and columns `cs` (zero-based) of an array. */
function pick(a: Arr, rs: readonly number[], cs: readonly number[]): Arr {
  allot(rs.length * cs.length)
  return selectArray(a, rs, cs)
}

const range = (from: number, to: number) => Array.from({ length: Math.max(0, to - from) }, (_, i) => from + i)

/** CHOOSECOLS/CHOOSEROWS indexes: numbers (an array gives each of its values), negative from the end, 0 or past the end #VALUE!. */
function indexes(segs: readonly Segment[], n: number, rules: Rules): number[] | CErr {
  const out: number[] = []
  for (const s of segs)
    for (const v of Array.isArray(s) ? (s as readonly V[]) : [s as V]) {
      const k = rules.toNum(v)
      if (isErr(k)) return k
      const i = Math.trunc(k)
      if (i === 0 || Math.abs(i) > n) return E.VALUE
      out.push(i > 0 ? i - 1 : n + i)
    }
  return out
}

/** TAKE and DROP: the first (last, when negative) `n` of `size`, as the zero-based span [from, to): a reference needs only its ends. */
function part(size: number, n: number | undefined, take: boolean): [from: number, to: number] {
  if (n === undefined) return [0, size]
  const k = Math.min(Math.abs(Math.trunc(n)), size)
  if (take) return n >= 0 ? [0, k] : [size - k, size]
  return n >= 0 ? [k, size] : [0, size - k]
}

/** TAKE and DROP: of a reference a reference (COUNTA skips its blank cells), of an array an array. */
function takeDrop(take: boolean): FnDef {
  return {
    args: 'R n n?',
    cx: true,
    array: true,
    fn: (cx: Cx, x: Arg, rows?: number, cols?: number) => {
      if (isErr(x)) return x
      if (take && ((rows !== undefined && Math.trunc(rows) === 0) || (cols !== undefined && Math.trunc(cols) === 0))) return EMPTY
      const g = cx.grid(x)
      const [r0, r1] = part(g.rows, rows, take)
      const [c0, c1] = part(g.cols, cols, take)
      if (r1 <= r0 || c1 <= c0) return EMPTY
      if (x instanceof Rng) return new Rng(x.s, x.r1 + r0, x.c1 + c0, x.r1 + r1 - 1, x.c1 + c1 - 1)
      return pick(fromGrid(g), range(r0, r1), range(c0, c1))
    },
  }
}

/** An argument of HSTACK, VSTACK as arrays: a value, a reference or an array; each area of a union or of a reference across sheets. */
const stackParts = (cx: Cx, x: Arg): Arr[] =>
  x instanceof Arr ? [x] : x instanceof Rng ? [fromGrid(cx.grid(x))] : x instanceof Areas ? x.list.flatMap((g) => stackParts(cx, g)) : [asArray(x)]

/**
 * Arrays side by side (`across`) or one under another; shorter ones are padded with #N/A. An error given as a
 * value is the answer (the first in order); a reference to an error cell is stacked, as in Excel.
 */
function stack(across: boolean): FnDef {
  return {
    args: 'P',
    rep: 1,
    cx: true,
    array: true,
    fn: (cx: Cx, xs: Arg[]) => {
      const e = xs.find(isErr)
      if (e) return e
      const parts = xs.flatMap((x) => stackParts(cx, x))
      const rows = across ? Math.max(...parts.map((p) => p.rows)) : parts.reduce((s, p) => s + p.rows, 0)
      const cols = across ? parts.reduce((s, p) => s + p.cols, 0) : Math.max(...parts.map((p) => p.cols))
      allot(rows * cols)
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

/** The values of a grid row by row (column by column), without blanks (ignore 1), errors (2) or both (3). */
function values(g: Grid, ignore: number, byCol: boolean, v: V[]) {
  const a = fromGrid(g)
  const [outer, inner] = byCol ? [a.cols, a.rows] : [a.rows, a.cols]
  // a whole column without blanks: only its occupied part
  const stop = ignore & 1 && a.fill === null ? (byCol ? Math.min(outer, a.w) : Math.min(outer, a.h)) : outer
  allot(v.length + stop * inner)
  for (let i = 0; i < stop; i++) {
    spend(inner)
    for (let j = 0; j < inner; j++) {
      const x = byCol ? a.at(j, i) : a.at(i, j)
      if ((x === null && ignore & 1) || (isErr(x) && ignore & 2)) continue
      v.push(x)
    }
  }
}

/** TOCOL, TOROW: the values of an array, or of each area of a union or of a reference across sheets in turn. */
function flatten(column: boolean): FnDef {
  return {
    args: 'R i? b?',
    cx: true,
    areas: true,
    array: true,
    fn: (cx: Cx, x: Arg, ignore = 0, byCol = false) => {
      if (isErr(x)) return x
      if (ignore < 0 || ignore > 3) return E.VALUE
      const v: V[] = []
      for (const part of x instanceof Areas ? x.list : [x]) values(cx.grid(part), ignore, byCol, v)
      if (!v.length) return EMPTY
      return column ? new Arr(v.length, 1, v.length, 1, v) : new Arr(1, v.length, 1, v.length, v)
    },
  }
}

/**
 * WRAPROWS, WRAPCOLS: a row or column cut into rows (columns) of `n`, the last padded (#N/A by default); fewer values
 * than `n` make one line of their own length, not padded (`WRAPROWS({1,2,3},5)` is {1,2,3}).
 */
function wrap(rows: boolean): FnDef {
  return {
    args: 'G n a?',
    array: true,
    fn: (g: Grid, count: number, pad: V = E.NA) => {
      if (g.rows > 1 && g.cols > 1) return E.VALUE
      if (Math.trunc(count) < 1) return E.NUM
      const a = fromGrid(g)
      const len = a.rows * a.cols
      const n = Math.min(Math.trunc(count), len)
      const lines = Math.ceil(len / n)
      allot(lines * n)
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
  CHOOSECOLS: { args: 'G V', rep: 1, max: NEWER_MAX_ARGS, array: true, cx: true, fn: (cx: Cx, g: Grid, segs: Segment[]) => {
    const cs = indexes(segs, g.cols, cx.rules)
    return isErr(cs) ? cs : pick(fromGrid(g), range(0, g.rows), cs)
  } },
  CHOOSEROWS: { args: 'G V', rep: 1, max: NEWER_MAX_ARGS, array: true, cx: true, fn: (cx: Cx, g: Grid, segs: Segment[]) => {
    const rs = indexes(segs, g.rows, cx.rules)
    return isErr(rs) ? rs : pick(fromGrid(g), rs, range(0, g.cols))
  } },
  DROP: takeDrop(false),
  EXPAND: {
    args: 'G n n? a?',
    array: true,
    fn: (g: Grid, rows?: number, cols?: number, pad: V = E.NA) => {
      const r = rows === undefined ? g.rows : Math.trunc(rows)
      const c = cols === undefined ? g.cols : Math.trunc(cols)
      if (r < g.rows || c < g.cols) return E.VALUE
      allot(r * c)
      const a = fromGrid(g)
      const v: V[] = []
      for (let i = 0; i < r; i++) for (let j = 0; j < c; j++) v.push(i < a.rows && j < a.cols ? a.at(i, j) : pad)
      return new Arr(r, c, r, c, v)
    },
  },
  HSTACK: { ...stack(true), max: NEWER_MAX_ARGS },
  TAKE: takeDrop(true),
  TOCOL: flatten(true),
  TOROW: flatten(false),
  // a value or a one-cell reference is that value (TYPE(TRANSPOSE(5)) is 1), an array of one value stays an array
  TRANSPOSE: { args: 'R', cx: true, array: true, fn: (cx: Cx, x: Arg) => (x instanceof Arr || isMany(x) ? transpose(fromGrid(cx.grid(x))) : cx.scalar(x)) },
  VSTACK: { ...stack(false), max: NEWER_MAX_ARGS },
  WRAPCOLS: wrap(false),
  WRAPROWS: wrap(true),
}
