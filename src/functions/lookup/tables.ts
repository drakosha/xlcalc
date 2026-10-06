/** VLOOKUP, HLOOKUP, MATCH, LOOKUP and INDEX, this engine's own (formula.js is case-sensitive, without wildcards or binary search). */
import { gridAt, type Arg, type Cx, type FnDef, type Grid } from '../../args.js'
import { Arr } from '../../arrays/value.js'
import { E, isErr, Rng, toBool, type V } from '../../values.js'
import { bisect, find, indexArg, keyOf, type Line } from './search.js'

/** Excel reads the index, looks up the key, and only then checks the index against the table. */
const tableLookup = (horizontal: boolean): FnDef => ({
  args: 'a G r a?',
  cx: true,
  fn: (cx: Cx, key: V, table: Grid, idx: V | Rng, approx: V = true): V => {
    const k = keyOf(key)
    if (isErr(k)) return k
    const n = indexArg(cx, idx)
    if (isErr(n)) return n
    const a = toBool(approx)
    if (isErr(a)) return a
    const pos = find(horizontal ? row(table) : column(table), k, a ? 1 : 0, true)
    if (pos < 0) return E.NA
    const i = Math.trunc(n)
    if (i < 1) return E.VALUE
    if (i > (horizontal ? table.rows : table.cols)) return E.REF
    return gridAt(table, horizontal ? (i - 1) * table.cols + pos : pos * table.cols + i - 1)
  },
})

const row = (g: Grid): Line => ({ at: (j) => gridAt(g, j), n: g.cols, filled: g.w })
const column = (g: Grid): Line => ({ at: (j) => gridAt(g, j * g.cols), n: g.rows, filled: g.h })

/** A wide range is read along its first row, any other along its first column. */
const along = (g: Grid): Line => (g.cols > g.rows ? row(g) : column(g))

function match(cx: Cx, key: V, range: Grid, type: V | Rng = 1): V {
  const k = keyOf(key)
  if (isErr(k)) return k
  const t = indexArg(cx, type, 'type')
  if (isErr(t)) return t
  if (range.rows !== 1 && range.cols !== 1) return E.NA
  // the type's sign counts, not its whole part: −0.5 is a descending match
  const pos = find(along(range), k, Math.sign(t), true)
  return pos < 0 ? E.NA : pos + 1
}

/**
 * LOOKUP: array form reads the last row (column) of the array, vector form the result at the same position.
 * Past the end of the result Excel reads further cells; here that is #REF!, since a session recomputes a
 * formula only for edits inside its references.
 */
function lookup(key: V, look: Grid, result?: Grid): V {
  const k = keyOf(key)
  if (isErr(k)) return k
  const pos = bisect(along(look), k)
  if (pos < 0) return E.NA
  if (result === undefined) return gridAt(look, look.cols > look.rows ? (look.rows - 1) * look.cols + pos : pos * look.cols + look.cols - 1)
  if (result.rows > 1 && result.cols > 1) return E.NA
  return pos < result.rows * result.cols ? gridAt(result, pos) : E.REF
}

/** Rows r1…r2 and columns c1…c2 of an array, as an array (a 1 × 1 one is its value). */
function slice(a: Arr, r1: number, c1: number, r2: number, c2: number): V | Arr {
  if (r1 === r2 && c1 === c2) return a.at(r1, c1)
  if (r1 === 0 && c1 === 0 && r2 === a.rows - 1 && c2 === a.cols - 1) return a
  const v: V[] = []
  for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) v.push(a.at(r, c))
  return new Arr(r2 - r1 + 1, c2 - c1 + 1, r2 - r1 + 1, c2 - c1 + 1, v)
}

/** INDEX: row or column 0 selects all of them; the answer is a reference (of a reference) or an array (of an array). */
function index(x: Arg, r: number, c?: number, area?: number): Arg {
  if (isErr(x)) return x
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
  return x instanceof Arr ? slice(x, r1, c1, r2, c2) : x
}

export const functions: Record<string, FnDef> = {
  HLOOKUP: tableLookup(true),
  INDEX: { args: 'R i i? i?', array: true, fn: index },
  LOOKUP: { args: 'a G G?', fn: lookup },
  MATCH: { args: 'a G r?', cx: true, fn: match },
  VLOOKUP: tableLookup(false),
}
