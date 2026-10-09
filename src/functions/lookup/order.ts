/**
 * SORT, SORTBY, UNIQUE and FILTER (Excel 365), this engine's own. Text compares case-insensitively; a sort is
 * stable; an empty answer is #CALC! in Excel, #VALUE! here (`EMPTY`). A whole column is handled by its
 * occupied rows: past them every row is blank.
 */
import type { Arg, Cx, FnDef, Grid } from '../../args.js'
import { Arr, asArray, EMPTY, fromGrid, selectArray } from '../../arrays/value.js'
import { allot } from '../../exec/budget.js'
import type { Rules } from '../../locale/rules.js'
import { Areas, E, isErr, Rng, type CErr, type V } from '../../values.js'
import { NEWER_MAX_ARGS } from '../common.js'

/** Sort rank of a type: numbers, text, booleans, errors, then blanks. */
const rank = (v: V) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : v === null ? 4 : 3)

/** Excel's sort order of two values, text by the book's locale; `desc` reverses it, blanks stay last. */
function order(a: V, b: V, desc: boolean, rules: Rules): number {
  const ra = rank(a)
  const rb = rank(b)
  if (ra === 4 || rb === 4) return ra - rb
  let c = ra - rb
  if (!c && ra === 0) c = (a as number) - (b as number)
  else if (!c && ra === 1) c = rules.compareText(a as string, b as string)
  else if (!c && ra === 2) c = Number(a) - Number(b)
  return desc ? -c : c
}

/** The lines (rows, or columns `byCol`) of an array that are not all blank past the occupied part, in order. */
const lineCount = (a: Arr, byCol: boolean) => (a.fill === null ? (byCol ? a.w : a.h) : byCol ? a.cols : a.rows)

/** Lines sorted by keys: `key(line, k)` is the k-th key value, `desc[k]` its direction; the rest of the lines (blank) last. */
function sortLines(a: Arr, byCol: boolean, key: (line: number, k: number) => V, desc: readonly boolean[], rules: Rules): Arr {
  const n = lineCount(a, byCol)
  const lines = Array.from({ length: n }, (_, i) => i).sort((x, y) => {
    for (let k = 0; k < desc.length; k++) {
      const c = order(key(x, k), key(y, k), desc[k], rules)
      if (c) return c
    }
    return x - y
  })
  // the sorted lines first, the blank ones past the occupied part after them
  const out = lineArr(a, lines, byCol)
  return byCol ? new Arr(a.rows, a.cols, a.rows, n, out.v, a.fill) : new Arr(a.rows, a.cols, n, a.cols, out.v, a.fill)
}

/** The given rows (columns `byCol`) of an array, in that order. */
function lineArr(a: Arr, lines: readonly number[], byCol: boolean): Arr {
  return selectArray(a, byCol ? undefined : lines, byCol ? lines : undefined)
}

/** Sort indexes (1…n) and orders (1 or −1): numbers, an array giving several. */
function keys(idx: Grid | undefined, ord: Grid | undefined, n: number, rules: Rules): { at: number[]; desc: boolean[] } | CErr {
  const nums = (g: Grid | undefined, d: number): number[] | CErr => {
    if (!g) return [d]
    const out: number[] = []
    for (const v of fromGrid(g).values()) {
      const x = rules.toNum(v)
      if (isErr(x)) return x
      out.push(Math.trunc(x))
    }
    return out
  }
  const at = nums(idx, 1)
  if (isErr(at)) return at
  const os = nums(ord, 1)
  if (isErr(os)) return os
  if (at.some((i) => i < 1 || i > n) || os.some((o) => o !== 1 && o !== -1)) return E.VALUE
  return { at: at.map((i) => i - 1), desc: at.map((_, k) => (os.length === 1 ? os[0] : (os[k] ?? 1)) === -1) }
}

function sort(cx: Cx, g: Grid, idx?: Grid, ord?: Grid, byCol = false): V | Arr {
  const a = fromGrid(g)
  const k = keys(idx, ord, byCol ? a.rows : a.cols, cx.rules)
  if (isErr(k)) return k
  return sortLines(a, byCol, (line, j) => (byCol ? a.at(k.at[j], line) : a.at(line, k.at[j])), k.desc, cx.rules)
}

/** An array, one value too, or a range of several cells: as a sort order it makes SORTBY answer by position. */
const several = (x: Arg | undefined): x is Rng | Arr => x instanceof Arr || (x instanceof Rng && (x.r2 > x.r1 || x.c2 > x.c1))

/** A sort order: −1 descending, 1 (TRUE too) or not given ascending, an error itself, anything else #VALUE!. */
function descending(cx: Cx, o: V | undefined): boolean | CErr {
  if (o === undefined) return false
  const n = cx.rules.toNum(o)
  return isErr(n) ? n : n === 1 || n === -1 ? n === -1 : E.VALUE
}

/** SORTBY: the array, then by-arrays each with its order; the orders one value each, or by position (`byPosition`). */
function sortby(cx: Cx, array: Arg, rest: (Arg | undefined)[]): V | Arr {
  return rest.some((x, i) => i % 2 === 1 && several(x)) ? byPosition(cx, array, rest) : sortOnce(cx, array, rest)
}

/**
 * SORTBY with one value in each order: by-arrays of one column (row) as long as the array's rows (columns), or one
 * value that keeps the order. The array and the by-arrays are read as the grid kind reads them (`args.ts`): a
 * reference across sheets first (#REF!), then an error given as a value, then a union (#VALUE!).
 */
function sortOnce(cx: Cx, array: Arg, rest: (Arg | undefined)[]): V | Arr {
  const args = [array, ...rest]
  if (args.some((x) => x instanceof Areas && x.cube)) return E.REF
  const error = args.find(isErr)
  if (error) return error
  if (args.some((x) => x instanceof Areas)) return E.VALUE
  const a = fromGrid(cx.grid(array))
  const by: Arr[] = []
  const desc: boolean[] = []
  for (let i = 0; i < rest.length; i += 2) {
    by.push(fromGrid(cx.grid(rest[i] ?? null)))
    const order = rest[i + 1]
    const d = descending(cx, order === undefined ? undefined : cx.scalar(order))
    if (isErr(d)) return d
    desc.push(d)
  }
  const one = (b: Arr) => b.rows === 1 && b.cols === 1
  const byCol = by[0].rows === 1 && by[0].cols > 1
  for (const b of by) if (!one(b) && (byCol ? b.rows !== 1 || b.cols !== a.cols : b.cols !== 1 || b.rows !== a.rows)) return E.VALUE
  return sortLines(a, byCol, (line, k) => (one(by[k]) ? by[k].at(0, 0) : byCol ? by[k].at(0, line) : by[k].at(line, 0)), desc, cx.rules)
}

/**
 * SORTBY with an array or a range of several cells among its orders, as Excel answers it (the set `known-held`): an
 * array as large as the orders, where an order missing at a position (no row or column is repeated), an error or a
 * value other than 1 and −1 answers its error or #VALUE!, the first position the first value of the sort by the
 * orders there, and every other position blank.
 */
function byPosition(cx: Cx, array: Arg, rest: (Arg | undefined)[]): Arr {
  const orders = rest.map((x, i) => (i % 2 === 1 && several(x) ? (cx.view(x) as Arr) : x))
  const arrays = orders.filter((x, i): x is Arr => i % 2 === 1 && x instanceof Arr)
  const rows = Math.max(...arrays.map((o) => o.rows))
  const cols = Math.max(...arrays.map((o) => o.cols))
  // the first order at a position that is not 1 or −1 answers its error there
  const check = (r: number, c: number): CErr | undefined => {
    for (let i = 1; i < orders.length; i += 2) {
      const o = orders[i]
      const v = o instanceof Arr ? (r < o.rows && c < o.cols ? o.at(r, c) : E.VALUE) : o instanceof Areas ? (o.cube ? E.REF : E.VALUE) : o === undefined ? undefined : cx.scalar(o)
      const d = descending(cx, v)
      if (isErr(d)) return d
    }
  }
  // past the occupied part of every order (or past its size) each position is #VALUE! or a scalar order's error: one fill
  const h = Math.max(1, ...arrays.map((o) => (o.fill === null ? o.h : o.rows)))
  const w = Math.max(1, ...arrays.map((o) => (o.fill === null ? o.w : o.cols)))
  allot(h * w)
  const out: V[] = new Array(h * w)
  for (let r = 0; r < h; r++) for (let c = 0; c < w; c++) out[r * w + c] = check(r, c) ?? null
  if (out[0] === null) {
    const first = sortOnce(cx, array, orders.map((o, i) => (i % 2 === 1 && o instanceof Arr ? o.at(0, 0) : o)))
    out[0] = first instanceof Arr ? first.at(0, 0) : first
  }
  return new Arr(rows, cols, h, w, out, h < rows || w < cols ? (check(rows, cols) ?? null) : null)
}

/** A line's identity for UNIQUE: values by type, text in any case. */
function lineKey(a: Arr, line: number, byCol: boolean): string {
  const n = byCol ? a.rows : a.cols
  let key = ''
  for (let i = 0; i < n; i++) {
    const v = byCol ? a.at(i, line) : a.at(line, i)
    key += (v === null ? 'b' : isErr(v) ? `e${v.code}` : typeof v === 'string' ? `t${v.toLowerCase()}` : `${typeof v}${String(v)}`) + '\u0000'
  }
  return key
}

function unique(g: Grid, byCol = false, once = false): V | Arr {
  const a = fromGrid(g)
  const total = byCol ? a.cols : a.rows
  const n = lineCount(a, byCol)
  const seen = new Map<string, { line: number; count: number }>()
  // the lines past the occupied part are all the same blank line
  for (let i = 0; i < Math.min(total, n + 1); i++) {
    const k = lineKey(a, i, byCol)
    const s = seen.get(k)
    const count = i === n ? total - n : 1
    if (s) s.count += count
    else seen.set(k, { line: i, count })
  }
  const lines = [...seen.values()].filter((s) => !once || s.count === 1).map((s) => s.line)
  return lines.length ? lineArr(a, lines, byCol) : EMPTY
}

/**
 * FILTER: an include column (row) as long as the array's rows (columns); numbers and booleans, text #VALUE!,
 * an error is the answer. An include array is read lazily: past its occupied part only its fill counts.
 */
function filter(cx: Cx, g: Grid, inc: Arg, empty?: V | Arr): V | Arr {
  if (isErr(inc)) return inc
  const a = fromGrid(g)
  const f = asArray(cx.view(inc))
  const byCol = !(f.cols === 1 && f.rows === a.rows) && f.rows === 1 && f.cols === a.cols
  if (!byCol && !(f.cols === 1 && f.rows === a.rows)) return E.VALUE
  const kept: number[] = []
  const n = byCol ? a.cols : a.rows
  const limit = f.fill === null || f.fill === false || f.fill === 0 ? (byCol ? f.w : f.h) : n
  for (let i = 0; i < limit; i++) {
    const v = byCol ? f.at(0, i) : f.at(i, 0)
    if (isErr(v)) return v
    if (typeof v === 'string') return E.VALUE
    const b = cx.rules.toBool(v)
    if (isErr(b)) return b
    if (b) kept.push(i)
  }
  if (!kept.length) return empty === undefined ? EMPTY : empty
  return lineArr(a, kept, byCol)
}

export const functions: Record<string, FnDef> = {
  FILTER: { args: 'G R p?', cx: true, array: true, fn: filter },
  SORT: { args: 'G G? G? b?', array: true, cx: true, fn: sort },
  // each further by-array with its order: 2, 3, 5, 7… arguments (Excel does not load SORTBY with 4); the arguments as given, so orders of several values come before a union or a reference across sheets
  SORTBY: { args: 'R R p?', rep: 2, max: NEWER_MAX_ARGS, array: true, cx: true, areas: true, fn: sortby },
  UNIQUE: { args: 'G b? b?', array: true, fn: unique },
}
