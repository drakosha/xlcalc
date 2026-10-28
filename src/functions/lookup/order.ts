/**
 * SORT, SORTBY, UNIQUE and FILTER (Excel 365), this engine's own. Text compares case-insensitively; a sort is
 * stable; an empty answer is #CALC! in Excel, #VALUE! here (`EMPTY`). A whole column is handled by its
 * occupied rows: past them every row is blank.
 */
import type { Arg, Cx, FnDef, Grid } from '../../args.js'
import { isArea, lifts } from '../../arrays/lift.js'
import { Arr, asArray, EMPTY, fromGrid, selectArray } from '../../arrays/value.js'
import { allot } from '../../exec/budget.js'
import type { Rules } from '../../locale/rules.js'
import { Areas, E, isErr, Rng, type CErr, type V } from '../../values.js'
import { NEWER_MAX_ARGS } from '../common.js'
import { sortOrder } from '../../sort-order.js'

/** The lines (rows, or columns `byCol`) of an array that are not all blank past the occupied part, in order. */
const lineCount = (a: Arr, byCol: boolean) => (a.fill === null ? (byCol ? a.w : a.h) : byCol ? a.cols : a.rows)

/** Lines sorted by keys: `key(line, k)` is the k-th key value, `desc[k]` its direction; the rest of the lines (blank) last. */
function sortLines(a: Arr, byCol: boolean, key: (line: number, k: number) => V, desc: readonly boolean[], rules: Rules): Arr {
  const n = lineCount(a, byCol)
  const lines = (desc.length === 1 && numberSort(n, (line) => key(line, 0), desc[0])) || Array.from({ length: n }, (_, i) => i).sort((x, y) => {
    for (let k = 0; k < desc.length; k++) {
      const c = sortOrder(key(x, k), key(y, k), desc[k], rules.compare, true)
      if (c) return c
    }
    return x - y
  })
  // the sorted lines first, the blank ones past the occupied part after them
  const out = lineArr(a, lines, byCol)
  return byCol ? new Arr(a.rows, a.cols, a.rows, n, out.v, a.fill) : new Arr(a.rows, a.cols, n, a.cols, out.v, a.fill)
}

/**
 * The stable order of `n` lines by one key that is a number in every line, or undefined for another key: a radix sort
 * over the numbers' bits (16 at a time), the order the comparator sort gives (-0 equal to 0, ties in line order) at a
 * fraction of its cost on a long column. Short ones keep the comparator.
 */
function numberSort(n: number, key: (line: number) => V, desc: boolean): number[] | undefined {
  if (n < RADIX_LINES) return undefined
  const value = new Float64Array(n)
  for (let i = 0; i < n; i++) {
    const v = key(i)
    if (typeof v !== 'number') return undefined
    // -0 sorts as 0
    value[i] = v + 0
  }
  // the bits as unsigned 16-bit digits (little-endian, as every JS host) that order as the numbers do, reversed when
  // descending; least significant first
  const bits = new Uint16Array(value.buffer)
  for (let i = 0; i < n; i++) {
    const neg = bits[i * 4 + 3] & 0x8000
    for (let d = 0; d < 4; d++) {
      let x = neg ? ~bits[i * 4 + d] & 0xffff : d === 3 ? bits[i * 4 + d] ^ 0x8000 : bits[i * 4 + d]
      if (desc) x = ~x & 0xffff
      bits[i * 4 + d] = x
    }
  }
  let order = new Uint32Array(n)
  for (let i = 0; i < n; i++) order[i] = i
  let next = new Uint32Array(n)
  const count = new Uint32Array(0x10000)
  for (let d = 0; d < 4; d++) {
    count.fill(0)
    for (let i = 0; i < n; i++) count[bits[i * 4 + d]]++
    // a digit the same in every line moves nothing
    if (count[bits[d]] === n) continue
    for (let x = 0, at = 0; x < 0x10000; x++) {
      const c = count[x]
      count[x] = at
      at += c
    }
    for (let i = 0; i < n; i++) {
      const line = order[i]
      next[count[bits[line * 4 + d]]++] = line
    }
    ;[order, next] = [next, order]
  }
  return Array.from(order)
}
const RADIX_LINES = 256

/** The given rows (columns `byCol`) of an array, in that order; within `limits.cells` (FILTER can keep every row of a whole column). */
function lineArr(a: Arr, lines: readonly number[], byCol: boolean): Arr {
  allot(lines.length * (byCol ? a.rows : a.cols))
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

/** A sort order: −1 descending, 1 (TRUE too) or not given ascending, an error itself, anything else #VALUE!. */
function descending(cx: Cx, o: V | undefined): boolean | CErr {
  if (o === undefined) return false
  const n = cx.rules.toNum(o)
  return isErr(n) ? n : n === 1 || n === -1 ? n === -1 : E.VALUE
}

/**
 * SORTBY: the array, then by-arrays each with its order; the orders one value each, or by position (`byPosition`) where
 * an order is an array, one value too, or a range of several cells (what `lifts`).
 */
function sortby(cx: Cx, array: Arg, rest: (Arg | undefined)[]): V | Arr {
  return rest.some((x, i) => i % 2 === 1 && lifts(x)) ? byPosition(cx, array, rest) : sortOnce(cx, array, rest)
}

/**
 * SORTBY with one value in each order: by-arrays of one column (row) as long as the array's rows (columns), the
 * first by-array deciding which, as Excel answers (the set `sortby-one-held`); a by-array of one value fits an array
 * of one row or one column. The array and the by-arrays are read as the grid kind reads them (`args.ts`): a
 * reference across sheets first (#REF!), then an error given as a value, then a union (#VALUE!).
 */
function sortOnce(cx: Cx, array: Arg, rest: (Arg | undefined)[]): V | Arr {
  const s = sortKeys(cx, array, rest)
  return isErr(s) ? s : sortLines(s.a, s.byCol, s.key, s.desc, cx.rules)
}

/** What SORTBY sorts: the array, by lines (columns `byCol`), by the keys of each line in their directions. */
interface Keys {
  a: Arr
  byCol: boolean
  key: (line: number, k: number) => V
  desc: boolean[]
}

/** What SORTBY with one value in each order sorts by, or the error it answers, without sorting; `grid` reads an argument. */
function sortKeys(cx: Cx, array: Arg, rest: (Arg | undefined)[], grid = (x: Arg) => cx.grid(x)): Keys | CErr {
  const args = [array, ...rest]
  if (args.some((x) => x instanceof Areas && x.cube)) return E.REF
  const error = args.find(isErr)
  if (error) return error
  if (args.some((x) => x instanceof Areas)) return E.VALUE
  const a = fromGrid(grid(array))
  const by: Arr[] = []
  const desc: boolean[] = []
  for (let i = 0; i < rest.length; i += 2) {
    by.push(fromGrid(grid(rest[i] ?? null)))
    const order = rest[i + 1]
    const d = descending(cx, order === undefined ? undefined : cx.scalar(order))
    if (isErr(d)) return d
    desc.push(d)
  }
  // one value is one row and one column: it fits the rows of an array of one row, the columns of one of one column
  const fits = (b: Arr, byCol: boolean) => (byCol ? b.rows === 1 && b.cols === a.cols : b.cols === 1 && b.rows === a.rows)
  const byCol = !fits(by[0], false)
  if (!by.every((b) => fits(b, byCol))) return E.VALUE
  return { a, byCol, key: (line, k) => (byCol ? by[k].at(0, line) : by[k].at(line, 0)), desc }
}

/**
 * SORTBY with an array or a range of several cells among its orders, as Excel answers it (the sets `known-held`,
 * `lookup-held`, `lookup-rules`): an array as large as the orders, lifted as an operator lifts (an order of one row or
 * column repeats, past a shorter one #N/A). Each position, in rows, answers the error of the call with the orders
 * there; the first that sorts gives its first sorted value, the others a blank, or #VALUE! with more than one by-array.
 */
function byPosition(cx: Cx, array: Arg, rest: (Arg | undefined)[]): Arr {
  const orders = rest.map((x, i) => (i % 2 === 1 && lifts(x) ? (cx.view(x) as Arr) : x))
  const arrays = orders.filter((x, i): x is Arr => i % 2 === 1 && x instanceof Arr)
  const rows = Math.max(...arrays.map((o) => o.rows))
  const cols = Math.max(...arrays.map((o) => o.cols))
  // an order array's element at a position: its one row or column repeats, past a shorter one #N/A
  const element = (o: Arr, r: number, c: number): V => {
    const [i, j] = [o.rows === 1 ? 0 : r, o.cols === 1 ? 0 : c]
    return i < o.rows && j < o.cols ? o.at(i, j) : E.NA
  }
  const at = (r: number, c: number) => orders.map((o, i) => (i % 2 === 1 && o instanceof Arr ? element(o, r, c) : o))
  // the array and the by-arrays are read once for all positions
  const grids = new Map<Arg, Grid>()
  const grid = (x: Arg) => grids.get(x) ?? grids.set(x, cx.grid(x)).get(x)!
  let sorted = false
  const answer = (r: number, c: number): V => {
    const s = sortKeys(cx, array, at(r, c), grid)
    if (isErr(s)) return s
    if (sorted) return rest.length > 2 ? E.VALUE : null
    sorted = true
    return sortLines(s.a, s.byCol, s.key, s.desc, cx.rules).at(0, 0)
  }
  // past the occupied part of every order each position answers alike: one fill
  const h = Math.max(1, ...arrays.map((o) => (o.fill === null ? o.h : o.rows)))
  const w = Math.max(1, ...arrays.map((o) => (o.fill === null ? o.w : o.cols)))
  allot(h * w)
  const out: V[] = new Array(h * w)
  for (let r = 0; r < h; r++) for (let c = 0; c < w; c++) out[r * w + c] = answer(r, c)
  sorted = true
  return new Arr(rows, cols, h, w, out, h < rows || w < cols ? answer(h < rows ? h : 0, h < rows ? 0 : w) : null)
}

/**
 * Text's identity for UNIQUE: in any case, and one for the spellings Unicode makes equivalent (é as one code point or as
 * e and an accent, the ligature ﬁ and fi), as Excel merges them; a soft hyphen still tells two texts apart there (the
 * set `low-values-held`).
 */
const textKey = (v: string) => (/[^\x00-\x7f]/.test(v) ? v.normalize('NFKC') : v).toLowerCase()

/** A line's identity for UNIQUE: values by type, text by `textKey`. */
function lineKey(a: Arr, line: number, byCol: boolean): string {
  const n = byCol ? a.rows : a.cols
  let key = ''
  for (let i = 0; i < n; i++) {
    const v = byCol ? a.at(i, line) : a.at(line, i)
    key += (v === null ? 'b' : isErr(v) ? `e${v.code}` : typeof v === 'string' ? `t${textKey(v)}` : `${typeof v}${String(v)}`) + '\u0000'
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
function filter(cx: Cx, g: Grid, inc: Arg, empty?: V | Arr | Rng): V | Arr {
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
  // if_empty is answered as a value, a reference by its values (ISREF of it is FALSE)
  if (!kept.length) return empty === undefined ? EMPTY : empty instanceof Rng ? (isArea(empty) ? cx.view(empty) : cx.scalar(empty)) : empty
  return lineArr(a, kept, byCol)
}

export const functions: Record<string, FnDef> = {
  FILTER: { args: 'G R p?', cx: true, array: true, fn: filter },
  SORT: { args: 'G G? G? b?', array: true, cx: true, fn: sort },
  // each further by-array with its order: 2, 3, 5, 7… arguments (Excel does not load SORTBY with 4); the arguments as given, so orders of several values come before a union or a reference across sheets
  SORTBY: { args: 'R R p?', rep: 2, max: NEWER_MAX_ARGS, array: true, cx: true, areas: true, fn: sortby },
  UNIQUE: { args: 'G b? b?', array: true, fn: unique },
}
