/** VLOOKUP, HLOOKUP, MATCH, LOOKUP and INDEX, this engine's own (formula.js is case-sensitive, without wildcards or binary search). */
import { gridAt, type Cx, type FnDef, type Grid } from '../../args.js'
import { MAX_COLS, MAX_ROWS } from '../../model.js'
import { E, isErr, Rng, toBool, type V } from '../../values.js'
import { bisect, find, indexArg, keyOf, type Line } from './search.js'

/** Excel reads the index, looks up the key, and only then checks the index against the table. */
const tableLookup = (horizontal: boolean): FnDef => ({
  args: 'a G R a?',
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
  const t = indexArg(cx, type)
  if (isErr(t)) return t
  if (range.rows !== 1 && range.cols !== 1) return E.NA
  // the type's sign counts, not its whole part: −0.5 is a descending match
  const pos = find(along(range), k, Math.sign(t), true)
  return pos < 0 ? E.NA : pos + 1
}

/**
 * LOOKUP: array form reads the last row (column) of the array; vector form reads the result at the same
 * position, past its end too, as Excel extends a result reference down a column or along a row.
 */
function lookup(cx: Cx, key: V, look: Grid, result?: V | Rng): V {
  const k = keyOf(key)
  if (isErr(k)) return k
  const pos = bisect(along(look), k)
  if (pos < 0) return E.NA
  if (result === undefined) return gridAt(look, look.cols > look.rows ? (look.rows - 1) * look.cols + pos : pos * look.cols + look.cols - 1)
  if (!(result instanceof Rng)) return pos === 0 ? result : E.NA
  if (result.r2 > result.r1 && result.c2 > result.c1) return E.NA
  const down = result.r2 > result.r1
  const r = result.r1 + (down ? pos : 0)
  const c = result.c1 + (down ? 0 : pos)
  return r >= MAX_ROWS || c >= MAX_COLS ? E.REF : cx.scalar(new Rng(result.s, r, c, r, c))
}

/** Row or column 0 selects all of them: more than one cell is an array answer (no spilling), #VALUE!. */
function index(g: Grid, r: number, c?: number, area?: number): V {
  if (area !== undefined && area !== 1) return area < 1 ? E.VALUE : E.REF
  if (c === undefined) {
    // one number indexes a row or a column; in a two-dimensional range Excel answers #REF!
    if (g.rows > 1 && g.cols > 1) return r < 0 ? E.VALUE : E.REF
    if (g.rows === 1) [r, c] = [1, r]
    else c = 1
  }
  if (r < 0 || c < 0) return E.VALUE
  if (r > g.rows || c > g.cols) return E.REF
  if ((r === 0 && g.rows > 1) || (c === 0 && g.cols > 1)) return E.VALUE
  return gridAt(g, ((r || 1) - 1) * g.cols + (c || 1) - 1)
}

export const functions: Record<string, FnDef> = {
  HLOOKUP: tableLookup(true),
  INDEX: { args: 'G i i? i?', fn: index },
  LOOKUP: { args: 'a G R?', cx: true, fn: lookup },
  MATCH: { args: 'a G R?', cx: true, fn: match },
  VLOOKUP: tableLookup(false),
}
