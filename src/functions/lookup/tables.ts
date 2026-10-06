/** VLOOKUP, HLOOKUP, MATCH, LOOKUP and INDEX, this engine's own (formula.js is case-sensitive, without wildcards or binary search). */
import { gridAt, type FnDef, type Grid } from '../../args.js'
import { E, isErr, toNum, type CErr, type V } from '../../values.js'
import { bisect, find, keyOf } from './search.js'

/** Excel answers #REF! for an error in a column index or match type. */
function indexArg(x: V): number | CErr {
  if (isErr(x)) return E.REF
  const n = toNum(x)
  return isErr(n) ? n : Math.trunc(n)
}

const tableLookup = (horizontal: boolean): FnDef => ({
  args: 'a G a b?',
  fn: (key: V, table: Grid, idx: V, approx = true): V => {
    const k = keyOf(key)
    if (isErr(k)) return k
    const i = indexArg(idx)
    if (isErr(i)) return i
    if (i < 1) return E.VALUE
    if (i > (horizontal ? table.rows : table.cols)) return E.REF
    const at = horizontal ? (j: number) => gridAt(table, j) : (j: number) => gridAt(table, j * table.cols)
    const pos = find(at, horizontal ? table.w : table.h, k, approx ? 1 : 0, true)
    if (pos < 0) return E.NA
    return gridAt(table, horizontal ? (i - 1) * table.cols + pos : pos * table.cols + i - 1)
  },
})

/** The values of a one-row or one-column range, else null. */
function line(g: Grid): { at: (i: number) => V; n: number } | null {
  return g.rows === 1 || g.cols === 1 ? { at: (i) => g.v[i], n: g.v.length } : null
}

function match(key: V, range: Grid, type: V = 1): V {
  const k = keyOf(key)
  if (isErr(k)) return k
  const t = indexArg(type)
  if (isErr(t)) return t
  const l = line(range)
  if (!l) return E.NA
  const pos = find(l.at, l.n, k, Math.sign(t), true)
  return pos < 0 ? E.NA : pos + 1
}

/** Array form: a wide array is searched along its first row, else its first column, and read in the last one. */
function lookup(key: V, look: Grid, result?: Grid): V {
  const k = keyOf(key)
  if (isErr(k)) return k
  if (!result) {
    const wide = look.cols > look.rows
    const at = wide ? (j: number) => gridAt(look, j) : (j: number) => gridAt(look, j * look.cols)
    const pos = bisect(at, wide ? look.w : look.h, k)
    if (pos < 0) return E.NA
    return gridAt(look, wide ? (look.rows - 1) * look.cols + pos : pos * look.cols + look.cols - 1)
  }
  const l = line(look)
  if (!l || !line(result)) return E.NA
  const pos = bisect(l.at, l.n, k)
  return pos < 0 || pos >= result.rows * result.cols ? E.NA : gridAt(result, pos)
}

/** Row or column 0 selects all of them: more than one cell is an array answer (no spilling), #VALUE!. */
function index(g: Grid, r: number, c?: number, area?: number): V {
  if (area !== undefined && area !== 1) return area < 1 ? E.VALUE : E.REF
  if (c === undefined) {
    if (g.rows === 1 && g.cols !== 1) [r, c] = [1, r]
    else c = g.cols === 1 ? 1 : 0
  }
  if (r < 0 || c < 0) return E.VALUE
  if (r > g.rows || c > g.cols) return E.REF
  if ((r === 0 && g.rows > 1) || (c === 0 && g.cols > 1)) return E.VALUE
  return gridAt(g, ((r || 1) - 1) * g.cols + (c || 1) - 1)
}

export const functions: Record<string, FnDef> = {
  HLOOKUP: tableLookup(true),
  INDEX: { args: 'G i i? i?', fn: index },
  LOOKUP: { args: 'a G G?', fn: lookup },
  MATCH: { args: 'a G a?', fn: match },
  VLOOKUP: tableLookup(false),
}
