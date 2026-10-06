/**
 * SORT, SORTBY, UNIQUE and FILTER (Excel 365), this engine's own. Text compares case-insensitively; a sort is
 * stable; an empty answer is #CALC! in Excel, #VALUE! here (`EMPTY`). A whole column is handled by its
 * occupied rows: past them every row is blank.
 */
import type { Arg, Cx, FnDef, Grid } from '../../args.js'
import { Arr, EMPTY, fromGrid } from '../../arrays/value.js'
import { collator } from '../../criteria.js'
import { E, isErr, toBool, toNum, type CErr, type V } from '../../values.js'

/** Sort rank of a type: numbers, text, booleans, errors, then blanks. */
const rank = (v: V) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : v === null ? 4 : 3)

/** Excel's sort order of two values; `desc` reverses it, blanks stay last. */
function order(a: V, b: V, desc: boolean): number {
  const ra = rank(a)
  const rb = rank(b)
  if (ra === 4 || rb === 4) return ra - rb
  let c = ra - rb
  if (!c && ra === 0) c = (a as number) - (b as number)
  else if (!c && ra === 1) c = collator.compare(a as string, b as string)
  else if (!c && ra === 2) c = Number(a) - Number(b)
  return desc ? -c : c
}

/** The lines (rows, or columns `byCol`) of an array that are not all blank past the occupied part, in order. */
const lineCount = (a: Arr, byCol: boolean) => (a.fill === null ? (byCol ? a.w : a.h) : byCol ? a.cols : a.rows)

/** Lines sorted by keys: `key(line, k)` is the k-th key value, `desc[k]` its direction; the rest of the lines (blank) last. */
function sortLines(a: Arr, byCol: boolean, key: (line: number, k: number) => V, desc: readonly boolean[]): Arr {
  const n = lineCount(a, byCol)
  const lines = Array.from({ length: n }, (_, i) => i).sort((x, y) => {
    for (let k = 0; k < desc.length; k++) {
      const c = order(key(x, k), key(y, k), desc[k])
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
  const v: V[] = []
  if (byCol) for (let r = 0; r < a.rows; r++) for (const c of lines) v.push(a.at(r, c))
  else for (const r of lines) for (let c = 0; c < a.cols; c++) v.push(a.at(r, c))
  return byCol ? new Arr(a.rows, lines.length, a.rows, lines.length, v) : new Arr(lines.length, a.cols, lines.length, a.cols, v)
}

/** Sort indexes (1…n) and orders (1 or −1): numbers, an array giving several. */
function keys(idx: Grid | undefined, ord: Grid | undefined, n: number): { at: number[]; desc: boolean[] } | CErr {
  const nums = (g: Grid | undefined, d: number): number[] | CErr => {
    if (!g) return [d]
    const out: number[] = []
    for (const v of fromGrid(g).values()) {
      const x = toNum(v)
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

function sort(g: Grid, idx?: Grid, ord?: Grid, byCol = false): V | Arr {
  const a = fromGrid(g)
  const k = keys(idx, ord, byCol ? a.rows : a.cols)
  if (isErr(k)) return k
  return sortLines(a, byCol, (line, j) => (byCol ? a.at(k.at[j], line) : a.at(line, k.at[j])), k.desc)
}

/** SORTBY: by-arrays of one column (row) as long as the array's rows (columns), each with an optional order. */
function sortby(g: Grid, rest: Grid[]): V | Arr {
  const a = fromGrid(g)
  const by: Arr[] = []
  const desc: boolean[] = []
  for (let i = 0; i < rest.length; i += 2) {
    by.push(fromGrid(rest[i]))
    const o = rest[i + 1] ? toNum(fromGrid(rest[i + 1]).at(0, 0)) : 1
    if (isErr(o)) return o
    if (o !== 1 && o !== -1) return E.VALUE
    desc.push(o === -1)
  }
  const byCol = by[0].rows === 1 && by[0].cols > 1
  for (const b of by) if (byCol ? b.rows !== 1 || b.cols !== a.cols : b.cols !== 1 || b.rows !== a.rows) return E.VALUE
  return sortLines(a, byCol, (line, k) => (byCol ? by[k].at(0, line) : by[k].at(line, 0)), desc)
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
  const v = cx.view(inc)
  const f = v instanceof Arr ? v : new Arr(1, 1, 1, 1, [v])
  const byCol = !(f.cols === 1 && f.rows === a.rows) && f.rows === 1 && f.cols === a.cols
  if (!byCol && !(f.cols === 1 && f.rows === a.rows)) return E.VALUE
  const kept: number[] = []
  const n = byCol ? a.cols : a.rows
  const limit = f.fill === null || f.fill === false || f.fill === 0 ? (byCol ? f.w : f.h) : n
  for (let i = 0; i < limit; i++) {
    const v = byCol ? f.at(0, i) : f.at(i, 0)
    if (isErr(v)) return v
    if (typeof v === 'string') return E.VALUE
    const b = toBool(v)
    if (isErr(b)) return b
    if (b) kept.push(i)
  }
  if (!kept.length) return empty === undefined ? EMPTY : empty
  return lineArr(a, kept, byCol)
}

export const functions: Record<string, FnDef> = {
  FILTER: { args: 'G R p?', cx: true, array: true, fn: filter },
  SORT: { args: 'G G? G? b?', array: true, fn: sort },
  // each further by-array with its order: 2, 3, 5, 7… arguments (Excel does not load SORTBY with 4)
  SORTBY: { args: 'G G G?', rep: 2, array: true, fn: sortby },
  UNIQUE: { args: 'G b? b?', array: true, fn: unique },
}
