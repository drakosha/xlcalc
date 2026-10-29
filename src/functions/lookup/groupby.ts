/**
 * GROUPBY and PIVOTBY (Excel 365), this engine's own after Microsoft's documentation: the data rows grouped by their
 * row fields (and by their column fields across), each group's values given to the function, with totals, sorting
 * and a filter. A group's values go to the function as a column; a function of two parameters (PERCENTOF) gets the
 * values of the total it is relative to as the second. Groups compare as UNIQUE compares (text in any case), sort as
 * SORT sorts; a cell the summary leaves empty is empty text.
 */
import { gridAt, type Arg, type Cx, type FnDef, type Grid } from '../../args.js'
import { arr, Arr } from '../../arrays/value.js'
import { allot } from '../../exec/budget.js'
import { callValue } from '../../exec/lambda.js'
import { sortOrder } from '../../sort-order.js'
import { E, isErr, Lambda, type CErr, type V } from '../../values.js'
import { one } from '../lambda.js'
import { valueKey } from './order.js'

/** The label of the totals, as Excel writes it. */
const TOTAL = 'Total'
/** An empty cell of the summary. */
const EMPTY_CELL = ''

/** The functions to aggregate with: one, or a row or a column of them, and their labels. */
interface Functions {
  fns: Lambda[]
  labels: V[]
}

/** A function's label: a built-in function's name, a LAMBDA's own word. */
const labelOf = (f: Lambda): V => f.name ?? 'LAMBDA'

/** The function argument: a function, or an array of functions in a row or a column; a second row or column labels them. */
function functionsOf(x: Arg): Functions | CErr {
  if (x instanceof Lambda) return { fns: [x], labels: [labelOf(x)] }
  if (!(x instanceof Arr)) return isErr(x) ? x : E.VALUE
  const at = (r: number, c: number) => x.at(r, c) as unknown
  const line = x.rows <= 2 && x.cols > 1 ? 'row' : x.cols <= 2 ? 'col' : undefined
  if (!line) return E.VALUE
  const n = line === 'row' ? x.cols : x.rows
  const fns: Lambda[] = []
  for (let i = 0; i < n; i++) {
    const f = line === 'row' ? at(0, i) : at(i, 0)
    if (!(f instanceof Lambda)) return E.VALUE
    fns.push(f)
  }
  const labelled = line === 'row' ? x.rows === 2 : x.cols === 2
  const labels = fns.map((f, i) => (labelled ? (line === 'row' ? x.at(1, i) : x.at(i, 1)) : labelOf(f)))
  return { fns, labels }
}

/** Whether a function takes the values relative to which it aggregates as a second argument (PERCENTOF). */
const twoArgs = (f: Lambda) => f.takes(2) && !f.takes(1)

/** A group of data rows: its keys from the first field to its level, and the rows (indexes into the data). */
interface Group {
  keys: V[]
  rows: number[]
  /** Its groups one field deeper, in order. */
  children: Group[]
}

/** A line of a summary's axis: a group of the last field (`item`), a subtotal of a shallower one, or the grand total. */
export interface Line {
  kind: 'item' | 'sub' | 'total'
  keys: V[]
  rows: number[]
  /** The line's parent group's rows, for a function relative to the parent's total (PIVOTBY's relative_to 3 and 4). */
  parent: number[]
}

/** How an axis is laid out: totals (`depth`, negative at the top), the fields' relationship, sorting. */
interface AxisOptions {
  depth: number
  table: boolean
  /** Sort keys: a field (0 from the first) or a measured column (`measure`), each with its direction. */
  sort: { at: number; desc: boolean }[]
}

/** The groups of the rows by the fields' values, nested field by field; the first value of a group spells its key. */
function groups(fields: Grid, start: number, rows: readonly number[], depth: number, k: number): Group[] {
  const out: Group[] = []
  const byKey = new Map<V, Group>()
  for (const r of rows) {
    const v = gridAt(fields, (start + r) * fields.cols + depth)
    const key = valueKey(v)
    let g = byKey.get(key)
    if (!g) {
      byKey.set(key, (g = { keys: [v], rows: [], children: [] }))
      out.push(g)
    }
    g.rows.push(r)
  }
  if (depth + 1 < k)
    for (const g of out) g.children = groups(fields, start, g.rows, depth + 1, k).map((c) => ({ ...c, keys: [...g.keys, ...c.keys] }))
  return out
}

/** The table relationship: each distinct row of the fields' values a group, all at one level. */
function flatGroups(fields: Grid, start: number, rows: readonly number[], k: number): Group[] {
  const out: Group[] = []
  const byKey = new Map<string, Group>()
  for (const r of rows) {
    const keys = Array.from({ length: k }, (_, c) => gridAt(fields, (start + r) * fields.cols + c))
    const key = keys.map((v) => String(valueKey(v))).join('\u0000')
    let g = byKey.get(key)
    if (!g) {
      byKey.set(key, (g = { keys, rows: [], children: [] }))
      out.push(g)
    }
    g.rows.push(r)
  }
  return out
}

/**
 * The lines of an axis, its groups sorted: by the sort keys at each level (a field's own value at its level, a measured
 * column by `measure` of the group's rows), the fields left as they are ascending after them; subtotals of the levels
 * the depth takes after (or before) their groups, the grand total last (or first).
 */
function lines(fields: Grid, start: number, rows: readonly number[], k: number, o: AxisOptions, measure: (rows: number[], col: number) => V, cx: Cx): Line[] {
  const top = o.table ? flatGroups(fields, start, rows, k) : groups(fields, start, rows, 0, k)
  const compare = cx.rules.compare
  const order = (list: Group[], level: number): Group[] => {
    const keys = [...o.sort.filter((s) => s.at >= k || o.table || s.at === level), ...Array.from({ length: k }, (_, i) => ({ at: i, desc: false }))]
    const measured = new Map<Group, V[]>()
    const at = (g: Group, s: { at: number }) => {
      if (s.at < k) return o.table ? g.keys[s.at] : g.keys[level]
      let m = measured.get(g)
      if (!m) measured.set(g, (m = []))
      return (m[s.at] ??= measure(g.rows, s.at - k))
    }
    return list
      .map((g, i) => ({ g, i }))
      .sort((x, y) => {
        for (const s of keys) {
          if (!o.table && s.at < k && s.at !== level) continue
          const c = sortOrder(at(x.g, s), at(y.g, s), s.desc, compare, true)
          if (c) return c
        }
        return x.i - y.i
      })
      .map((x) => x.g)
  }
  const subs = Math.abs(o.depth) - 1
  const atTop = o.depth < 0
  const out: Line[] = []
  const visit = (list: Group[], level: number, parent: number[]) => {
    for (const g of order(list, level)) {
      const leaf = o.table || level === k - 1
      const sub: Line = { kind: 'sub', keys: [...g.keys, ...new Array<V>(k - g.keys.length).fill(EMPTY_CELL)], rows: g.rows, parent }
      const subtotal = !leaf && level < subs
      if (subtotal && atTop) out.push(sub)
      if (leaf) out.push({ kind: 'item', keys: g.keys, rows: g.rows, parent })
      else visit(g.children, level + 1, g.rows)
      if (subtotal && !atTop) out.push(sub)
    }
  }
  const all = [...rows]
  const total: Line = { kind: 'total', keys: [TOTAL, ...new Array<V>(k - 1).fill(EMPTY_CELL)], rows: all, parent: all }
  if (o.depth && atTop) out.push(total)
  visit(top, 0, all)
  if (o.depth && !atTop) out.push(total)
  return out
}

/** The headers of the fields' data: whether its first row is headers, and whether the summary shows them. */
function headersOf(mode: number | undefined, values: Grid, levels: number): { has: boolean; show: boolean } | CErr {
  if (mode === undefined) {
    // headers when the first value is text above a number, shown when there are several levels of fields
    const has = typeof gridAt(values, 0) === 'string' && typeof gridAt(values, values.cols) === 'number'
    return { has, show: has && levels > 1 }
  }
  if (mode === 0 || mode === 1 || mode === 2 || mode === 3) return { has: mode === 1 || mode === 3, show: mode >= 2 }
  return E.VALUE
}

/** A whole-number option from its argument (undefined for one not given); a number not whole is truncated. */
const option = (n: number | undefined) => (n === undefined ? undefined : Math.trunc(n))

/** The sort keys of a sort order: 1-based column numbers, negative descending; 0 or past the columns #VALUE!. */
function sortKeys(g: Grid | undefined, columns: number): AxisOptions['sort'] | CErr {
  if (!g) return []
  const out: AxisOptions['sort'] = []
  for (let i = 0; i < g.rows * g.cols; i++) {
    const v = gridAt(g, i)
    if (isErr(v)) return v
    if (typeof v !== 'number') return E.VALUE
    const n = Math.trunc(v)
    if (n === 0 || Math.abs(n) > columns) return E.VALUE
    out.push({ at: Math.abs(n) - 1, desc: n < 0 })
  }
  return out
}

/** The data rows a filter keeps (all without one): a column of booleans or numbers as long as the data, else #VALUE!. */
function kept(filter: Grid | undefined, n: number, start: number): number[] | CErr {
  const all = Array.from({ length: n - start }, (_, i) => i)
  if (!filter) return all
  if (filter.cols !== 1 || (filter.rows !== n && filter.rows !== n - start)) return E.VALUE
  const from = filter.rows === n ? start : 0
  const out: number[] = []
  for (const i of all) {
    const v = gridAt(filter, from + i)
    if (isErr(v)) return v
    if (typeof v === 'string') return E.VALUE
    if (v === true || (typeof v === 'number' && v !== 0)) out.push(i)
  }
  return out
}

/** A column of a grid's data rows (from `start`) at `rows`, as an array. */
const column = (g: Grid, start: number, col: number, rows: readonly number[]): Arr => arr(rows.length, 1, rows.map((r) => gridAt(g, (start + r) * g.cols + col)))

/** The aggregation of each value column by each function: the cell of rows `rows`, relative to `total`'s rows. */
function aggregator(cx: Cx, values: Grid, start: number, fns: Functions) {
  const m = values.cols
  return (rows: number[], col: number, total: number[]): V => {
    const f = fns.fns[Math.floor(col / m)]
    const c = col % m
    const x = callValue(f, twoArgs(f) ? [column(values, start, c, rows), column(values, start, c, total)] : [column(values, start, c, rows)])
    return one(x, cx) ?? E.CALC
  }
}

/** The labels of the measured columns, function by function, each over the value columns: a value's header, a function's label. */
function measuredLabels(values: Grid, has: boolean, fns: Functions): { headers: V[]; labels: V[] } {
  const m = values.cols
  const headers = fns.fns.flatMap(() => Array.from({ length: m }, (_, c) => (has ? gridAt(values, c) : m > 1 ? `Value ${c + 1}` : 'Value')))
  const labels = fns.labels.flatMap((l) => new Array<V>(m).fill(l))
  return { headers, labels }
}

/** The headers of fields: their own first row, or made up as `Row Field 1`, …. */
const fieldHeaders = (fields: Grid, has: boolean, word: string): V[] => Array.from({ length: fields.cols }, (_, c) => (has ? gridAt(fields, c) : `${word} ${c + 1}`))

/** The depth of totals: left out, grand totals and subtotals of every level; past the fields' levels #VALUE!. */
function depthOf(n: number | undefined, levels: number, table: boolean): number | CErr {
  const d = n ?? (table ? 1 : levels)
  if (Math.abs(d) > levels || (table && Math.abs(d) > 1)) return E.VALUE
  return d
}

/** The summary's rows as an array, a row at a time. */
function rowsArray(out: V[][]): Arr {
  const cols = out[0].length
  allot(out.length * cols)
  return arr(out.length, cols, out.flat())
}

function groupby(cx: Cx, rowFields: Grid, values: Grid, fn: Arg, fieldHeadersMode?: number, totalDepth?: number, sortOrderGrid?: Grid, filterGrid?: Grid, relationship?: number): Arg {
  const fns = functionsOf(fn)
  if (isErr(fns)) return fns
  const n = rowFields.rows
  if (values.rows !== n) return E.VALUE
  const k = rowFields.cols
  const rel = option(relationship) ?? 0
  if (rel !== 0 && rel !== 1) return E.VALUE
  const table = rel === 1 && k > 1
  const heads = headersOf(option(fieldHeadersMode), values, k)
  if (isErr(heads)) return heads
  const start = heads.has ? 1 : 0
  const depth = depthOf(option(totalDepth), k, table)
  if (isErr(depth)) return depth
  const measuredCols = values.cols * fns.fns.length
  const sort = sortKeys(sortOrderGrid, k + measuredCols)
  if (isErr(sort)) return sort
  const rows = kept(filterGrid, n, start)
  if (isErr(rows)) return rows
  if (!rows.length) return E.CALC
  const agg = aggregator(cx, values, start, fns)
  const axis = lines(rowFields, start, rows, k, { depth, table, sort }, (r, c) => agg(r, c, rows), cx)
  const out: V[][] = []
  const { headers, labels } = measuredLabels(values, heads.has, fns)
  if (fns.fns.length > 1) out.push([...new Array<V>(k).fill(EMPTY_CELL), ...labels])
  if (heads.show) out.push([...fieldHeaders(rowFields, heads.has, 'Row Field'), ...headers])
  for (const line of axis) out.push([...line.keys, ...Array.from({ length: measuredCols }, (_, c) => agg(line.rows, c, rows))])
  return rowsArray(out)
}

/** The rows of `a` that are also in `b`. */
const meet = (a: readonly number[], b: ReadonlySet<number>) => a.filter((r) => b.has(r))

function pivotby(cx: Cx, rowFields: Grid, colFields: Grid, values: Grid, fn: Arg, fieldHeadersMode?: number, rowDepth?: number, rowSort?: Grid, colDepth?: number, colSort?: Grid, filterGrid?: Grid, relativeTo?: number): Arg {
  const fns = functionsOf(fn)
  if (isErr(fns)) return fns
  const n = rowFields.rows
  if (values.rows !== n || colFields.rows !== n) return E.VALUE
  const [k, j] = [rowFields.cols, colFields.cols]
  const heads = headersOf(option(fieldHeadersMode), values, Math.max(k, j))
  if (isErr(heads)) return heads
  const start = heads.has ? 1 : 0
  const rel = option(relativeTo) ?? 0
  if (rel < 0 || rel > 4) return E.VALUE
  const [rd, cd] = [depthOf(option(rowDepth), k, false), depthOf(option(colDepth), j, false)]
  if (isErr(rd)) return rd
  if (isErr(cd)) return cd
  const measuredCols = values.cols * fns.fns.length
  const [rs, cs] = [sortKeys(rowSort, k + measuredCols), sortKeys(colSort, j + measuredCols)]
  if (isErr(rs)) return rs
  if (isErr(cs)) return cs
  const rows = kept(filterGrid, n, start)
  if (isErr(rows)) return rows
  if (!rows.length) return E.CALC
  const agg = aggregator(cx, values, start, fns)
  const rowLines = lines(rowFields, start, rows, k, { depth: rd, table: false, sort: rs }, (r, c) => agg(r, c, rows), cx)
  const colLines = lines(colFields, start, rows, j, { depth: cd, table: false, sort: cs }, (r, c) => agg(r, c, rows), cx)
  const { headers, labels } = measuredLabels(values, heads.has, fns)
  const out: V[][] = []
  // a header row for each column field's level, then the labels of several measured columns
  for (let level = 0; level < j; level++) out.push([...new Array<V>(k).fill(EMPTY_CELL), ...colLines.flatMap((c) => new Array<V>(measuredCols).fill(c.keys[level]))])
  if (measuredCols > 1) out.push([...new Array<V>(k).fill(EMPTY_CELL), ...colLines.flatMap(() => (fns.fns.length > 1 ? labels : headers))])
  if (heads.show) out.push([...fieldHeaders(rowFields, heads.has, 'Row Field'), ...colLines.flatMap(() => new Array<V>(measuredCols).fill(EMPTY_CELL))])
  const sets = new Map<Line, Set<number>>()
  const setOf = (l: Line) => {
    let s = sets.get(l)
    if (!s) sets.set(l, (s = new Set(l.rows)))
    return s
  }
  for (const r of rowLines) {
    const cells: V[] = []
    for (const c of colLines) {
      const here = meet(r.rows, setOf(c))
      // the total a function of two parameters is relative to: the column's, the row's, the grand, a parent's
      const total = rel === 0 ? c.rows : rel === 1 ? r.rows : rel === 2 ? rows : rel === 3 ? meet(r.rows, new Set(c.parent)) : meet(c.rows, new Set(r.parent))
      for (let m = 0; m < measuredCols; m++) cells.push(here.length ? agg(here, m, total) : EMPTY_CELL)
    }
    out.push([...r.keys, ...cells])
  }
  return rowsArray(out)
}

/** The optional arguments left out between commas take their defaults. */
const defaults = (from: number, to: number) => ({ at: Object.fromEntries(Array.from({ length: to - from + 1 }, (_, i) => [from + i, 'default' as const])) })

export const functions: Record<string, FnDef> = {
  GROUPBY: { args: 'G G R n? n? G? G? n?', omit: defaults(3, 7), cx: true, array: true, fn: groupby },
  PIVOTBY: { args: 'G G G R n? n? G? n? G? G? n?', omit: defaults(4, 10), cx: true, array: true, fn: pivotby },
}
