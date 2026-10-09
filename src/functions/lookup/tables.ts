/** VLOOKUP, HLOOKUP, MATCH, LOOKUP and INDEX, this engine's own (formula.js is case-sensitive, without wildcards or binary search). */
import { gridAt, type Arg, type Cx, type FnDef, type Grid } from '../../args.js'
import { arr, Arr } from '../../arrays/value.js'
import type { Rules } from '../../locale/rules.js'
import { Areas, E, isErr, Rng, whole, type CErr, type V } from '../../values.js'
import { bisect, find, indexArg, keyOf, type Line } from './search.js'

/**
 * A table given as one value, not a reference or an array: Excel reads it as a number and finds nothing in it, so
 * `MATCH(5,5,0)` is #N/A and `MATCH(5,"x",0)` #VALUE! (the set `single-values`); null for a table to search.
 */
const oneValue = (cx: Cx, table: Arg): CErr | null => (table instanceof Rng || table instanceof Arr ? null : isErr(cx.rules.toNum(cx.scalar(table))) ? E.VALUE : E.NA)

/**
 * Excel reads the index, looks up the key, and only then checks the index against the table. An array or a
 * range of several cells as the index answers the found line at each number, shaped as the index; the first
 * number in order that is not valid is the whole answer; a call lifted by the key or the match mode reads the
 * index's top-left number (`liftFirst`).
 */
const tableLookup = (horizontal: boolean): FnDef => ({
  args: 'a R R a?',
  cx: true,
  array: true,
  liftFirst: [2],
  fn: (cx: Cx, key: V, given: Arg, idx: V | Rng | Arr, approx: V = true): V | Arr => {
    if (isErr(given)) return given
    const k = keyOf(key)
    if (isErr(k)) return k
    const many = idx instanceof Arr || (idx instanceof Rng && (idx.r2 > idx.r1 || idx.c2 > idx.c1))
    const n = many ? 0 : indexArg(cx, idx as V | Rng)
    if (isErr(n)) return n
    const a = cx.rules.toBool(approx)
    if (isErr(a)) return a
    const one = oneValue(cx, given)
    if (one) return one
    const table = cx.grid(given)
    const pos = find(horizontal ? row(table) : column(table), k, a ? 1 : 0, true, cx.rules.compare)
    if (pos < 0) return E.NA
    const size = horizontal ? table.rows : table.cols
    const bad = (i: number) => (i < 1 ? E.VALUE : i > size ? E.REF : undefined)
    const cell = (i: number) => gridAt(table, horizontal ? (i - 1) * table.cols + pos : pos * table.cols + i - 1)
    if (many) return lineAt(cx.view(idx) as Arr, bad, cell, cx.rules)
    const i = whole(n)
    return bad(i) ?? cell(i)
  },
})

/** The found line at every number of an index array, row by row; an error or text that is not a number is #REF!. */
function lineAt(idx: Arr, bad: (i: number) => CErr | undefined, cell: (i: number) => V, rules: Rules): CErr | Arr {
  const out: V[] = []
  for (let r = 0; r < idx.rows; r++)
    for (let c = 0; c < idx.cols; c++) {
      const n = rules.toNum(idx.at(r, c))
      if (isErr(n)) return E.REF
      const i = whole(n)
      const e = bad(i)
      if (e) return e
      out.push(cell(i))
    }
  return arr(idx.rows, idx.cols, out)
}

const row = (g: Grid): Line => ({ at: (j) => gridAt(g, j), n: g.cols, filled: g.w, values: g.v, stride: 1 })
const column = (g: Grid): Line => ({ at: (j) => gridAt(g, j * g.cols), n: g.rows, filled: g.h, values: g.v, stride: g.w })

/** A wide range is read along its first row, any other along its first column. */
const along = (g: Grid): Line => (g.cols > g.rows ? row(g) : column(g))

function match(cx: Cx, key: V, given: Arg, type: V | Rng = 1): V {
  if (isErr(given)) return given
  const k = keyOf(key)
  if (isErr(k)) return k
  const t = indexArg(cx, type, 'type')
  if (isErr(t)) return t
  const one = oneValue(cx, given)
  if (one) return one
  const range = cx.grid(given)
  if (range.rows !== 1 && range.cols !== 1) return E.NA
  // the type's sign counts, not its whole part: −0.5 is a descending match
  const pos = find(along(range), k, Math.sign(t), true, cx.rules.compare)
  return pos < 0 ? E.NA : pos + 1
}

/**
 * LOOKUP: array form reads the last row (column) of the array, vector form the result at the same position.
 * Excel reads on past the end of a shorter result range: the walk widens a result row or column to the lookup
 * range's length (`exec/walk.ts`), so those cells are among the formula's references; past what it widens, which a
 * session would not recompute for, it is #REF!. Past the end of a value or an array it is #N/A.
 */
function lookup(cx: Cx, key: V, look: Grid, result?: Arg): V {
  const k = keyOf(key)
  if (isErr(k)) return k
  if (isErr(result)) return result
  const pos = bisect(along(look), k, cx.rules.compare)
  if (pos < 0) return E.NA
  if (result === undefined) return gridAt(look, look.cols > look.rows ? (look.rows - 1) * look.cols + pos : pos * look.cols + look.cols - 1)
  const g = cx.grid(result)
  if (g.rows > 1 && g.cols > 1) return E.NA
  return pos < g.rows * g.cols ? gridAt(g, pos) : result instanceof Rng ? E.REF : E.NA
}

/** Rows r1…r2 and columns c1…c2 of an array, as an array (a 1 × 1 one is its value). */
function slice(a: Arr, r1: number, c1: number, r2: number, c2: number): Arr {
  if (r1 === 0 && c1 === 0 && r2 === a.rows - 1 && c2 === a.cols - 1) return a
  const v: V[] = []
  for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) v.push(a.at(r, c))
  return new Arr(r2 - r1 + 1, c2 - c1 + 1, r2 - r1 + 1, c2 - c1 + 1, v)
}

/**
 * INDEX: row or column 0 selects all of them; the answer is a reference (of a reference) or an array (of an array).
 * Of a union the area numbered `area` (the first by default) is indexed; a reference across sheets is #VALUE!.
 */
function index(x: Arg, r: number, c?: number, area?: number): Arg {
  if (isErr(x)) return x
  if (x instanceof Areas) {
    if (x.cube) return E.VALUE
    const n = area ?? 1
    if (n < 1) return E.VALUE
    if (n > x.list.length) return E.REF
    return index(x.list[n - 1], r, c)
  }
  const rows = x instanceof Rng ? x.r2 - x.r1 + 1 : x instanceof Arr ? x.rows : 1
  const cols = x instanceof Rng ? x.c2 - x.c1 + 1 : x instanceof Arr ? x.cols : 1
  if (area !== undefined && area !== 1) return area < 1 ? E.VALUE : E.REF
  if (c === undefined) {
    // one number indexes a row or a column; in a two-dimensional range Excel answers #REF!
    if (rows > 1 && cols > 1) return r < 0 ? E.VALUE : E.REF
    if (rows === 1) [r, c] = [1, r]
    else c = 1
  }
  if (r < 0 || c < 0) return E.VALUE
  if (r > rows || c > cols) return E.REF
  const [r1, r2] = r ? [r - 1, r - 1] : [0, rows - 1]
  const [c1, c2] = c ? [c - 1, c - 1] : [0, cols - 1]
  // a reference, one cell too: SUM(INDEX(A1:A3,2)) skips text there, as a range does
  if (x instanceof Rng) return new Rng(x.s, x.r1 + r1, x.c1 + c1, x.r1 + r2, x.c1 + c2)
  // an element of an array is a value; a line selected by 0 stays an array, 1 × 1 too (TYPE(INDEX({1;2},1,0)+0) is 64)
  if (x instanceof Arr) return r && c ? x.at(r1, c1) : slice(x, r1, c1, r2, c2)
  return x
}

export const functions: Record<string, FnDef> = {
  HLOOKUP: tableLookup(true),
  INDEX: { args: 'R w w? w?', array: true, areas: true, fn: index },
  LOOKUP: { args: 'a G R?', cx: true, fn: lookup },
  MATCH: { args: 'a R r?', cx: true, fn: match },
  VLOOKUP: tableLookup(false),
}
