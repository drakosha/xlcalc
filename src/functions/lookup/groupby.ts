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

/** The labels of the grand total, as Excel writes them: with subtotals beside it, and alone. */
const GRAND = 'Grand Total'
const TOTAL = 'Total'
/** An empty cell of the summary. */
const EMPTY_CELL = ''

/** The functions to aggregate with: one, or a row or a column of them, and their labels. */
interface Functions {
  fns: Lambda[]
  labels: V[]
  /** Given as a column: each group's line repeats for each function, its label beside it; else they go across. */
  down: boolean
}

/** A function's label: a built-in function's name, a LAMBDA's `CUSTOM`, as Excel writes them. */
const labelOf = (f: Lambda): V => f.name ?? 'CUSTOM'

/** The function argument: a function, or an array of functions in a row or a column; a second row or column labels them. */
function functionsOf(x: Arg): Functions | CErr {
  if (x instanceof Lambda) return { fns: [x], labels: [labelOf(x)], down: false }
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
  return { fns, labels, down: line === 'col' && n > 1 }
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

/**
 * The groups of the rows by the fields' values, nested field by field, in the order of `rows`: the first row of a group
 * there spells its key; a group's rows stay in the data's order, as the function gets them.
 */
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
  for (const g of out) {
    if (depth + 1 < k) g.children = groups(fields, start, g.rows, depth + 1, k).map((c) => ({ ...c, keys: [...g.keys, ...c.keys] }))
    g.rows.sort((x, y) => x - y)
  }
  return out
}

/** The groups of the last field under `list`, at one level (the table relationship). */
const leaves = (list: Group[]): Group[] => list.flatMap((g) => (g.children.length ? leaves(g.children) : [g]))

/**
 * The lines of an axis, its groups sorted: by the sort keys at each level (a field's own value at its level, a measured
 * column by `measure` of the group's rows), the fields left as they are ascending after them; subtotals of the levels
 * the depth takes after (or before) their groups, the grand total last (or first).
 */
function lines(fields: Grid, start: number, rows: readonly number[], k: number, o: AxisOptions, measure: (rows: number[], col: number) => V, cx: Cx): Line[] {
  const compare = cx.rules.compare
  // the rows in the order of their fields, as sorted (ascending where the sort leaves a field): a group's first row there
  // spells its key (`West` and `west` show as the one that sorts first)
  const field = (r: number, c: number) => gridAt(fields, (start + r) * fields.cols + c)
  const byFields = [...o.sort.filter((x) => x.at < k), ...Array.from({ length: k }, (_, at) => ({ at, desc: false }))]
  const inOrder = [...rows].sort((a, b) => {
    for (const { at, desc } of byFields) {
      const d = sortOrder(field(a, at), field(b, at), desc, compare, true)
      if (d) return d
    }
    return a - b
  })
  const nested = groups(fields, start, inOrder, 0, k)
  const top = o.table ? leaves(nested) : nested
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
  const total: Line = { kind: 'total', keys: [subs > 0 ? GRAND : TOTAL, ...new Array<V>(k - 1).fill(EMPTY_CELL)], rows: all, parent: all }
  if (o.depth && atTop) out.push(total)
  visit(top, 0, all)
  if (o.depth && !atTop) out.push(total)
  return out
}

/**
 * The headers of the fields' data: whether its first row is headers, and whether the summary shows them. Left out, as
 * Excel for Mac 16.114 answered the sets `groupby` and `groupby-held`: where the first row field and the first value are
 * text, the first row is headers, not shown, if every value under it is a number or blank, and the call is #VALUE!
 * otherwise; elsewhere there are none.
 */
function headersOf(mode: number | undefined, fields: Grid, values: Grid): { has: boolean; show: boolean } | CErr {
  if (mode === undefined) {
    if (typeof gridAt(fields, 0) !== 'string' || typeof gridAt(values, 0) !== 'string') return { has: false, show: false }
    for (let i = values.cols; i < values.rows * values.cols; i++) {
      const v = gridAt(values, i)
      if (v !== null && typeof v !== 'number') return E.VALUE
    }
    return { has: true, show: false }
  }
  if (mode === 0 || mode === 1 || mode === 2 || mode === 3) return { has: mode === 1 || mode === 3, show: mode >= 2 }
  return E.VALUE
}

/** A whole-number option from its argument (undefined for one not given); one not whole is NaN, which no option takes (#VALUE!). */
const option = (n: number | undefined) => (n === undefined ? undefined : Number.isInteger(n) ? n : NaN)

/**
 * The sort keys of a sort order: whole 1-based column numbers, negative descending; 0 or past the columns #VALUE!, and so are
 * several keys with a measured column (`fields` the fields' columns) among them, as the documentation takes several only
 * of the fields.
 */
function sortKeys(g: Grid | undefined, columns: number, fields: number): AxisOptions['sort'] | CErr {
  if (!g) return []
  const out: AxisOptions['sort'] = []
  for (let i = 0; i < g.rows * g.cols; i++) {
    const v = gridAt(g, i)
    if (isErr(v)) return v
    if (typeof v !== 'number') return E.VALUE
    const n = v
    if (!Number.isInteger(n) || n === 0 || Math.abs(n) > columns) return E.VALUE
    out.push({ at: Math.abs(n) - 1, desc: n < 0 })
  }
  return out.length > 1 && out.some((x) => x.at >= fields) ? E.VALUE : out
}

/** The data rows a filter keeps (all without one): a column of booleans or numbers as long as the data, else #VALUE!; none #VALUE! too. */
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
  return out.length ? out : E.VALUE
}

/** A column of a grid's data rows (from `start`) at `rows`, as an array. */
const column = (g: Grid, start: number, col: number, rows: readonly number[]): Arr => arr(rows.length, 1, rows.map((r) => gridAt(g, (start + r) * g.cols + col)))

/** A function's answer that is not one value: the whole summary is #CALC!. */
class NotOne extends Error {}

/** The aggregation of each value column by each function: the cell of rows `rows`, relative to `total`'s rows. */
function aggregator(cx: Cx, values: Grid, start: number, fns: Functions) {
  const m = values.cols
  return (rows: number[], col: number, total: number[]): V => {
    const f = fns.fns[Math.floor(col / m)]
    const c = col % m
    const x = callValue(f, twoArgs(f) ? [column(values, start, c, rows), column(values, start, c, total)] : [column(values, start, c, rows)])
    // an array, of one value too, is not one value here, as Excel for Mac answers `LAMBDA(x,TAKE(x,1))` (MAP takes it)
    const v = x instanceof Arr ? undefined : one(x, cx)
    if (v === undefined) throw new NotOne()
    return v
  }
}

/** A summary, or #CALC! when a function answered more than one value. */
function summary(make: () => Arg): Arg {
  try {
    return make()
  } catch (e) {
    if (e instanceof NotOne) return E.CALC
    throw e
  }
}

/** The labels of the measured columns, function by function, each over the value columns: a value's header (made up as `Value 1`, …), a function's label. */
function measuredLabels(values: Grid, has: boolean, fns: Functions, across: boolean): { headers: V[]; labels: V[] } {
  const m = values.cols
  const per = across ? fns.fns.length : 1
  const headers = Array.from({ length: per }, () => Array.from({ length: m }, (_, c) => (has ? gridAt(values, c) : `Value ${c + 1}`))).flat()
  const labels = fns.labels.flatMap((l) => new Array<V>(m).fill(l))
  return { headers, labels }
}

/** The headers of fields: their own first row, or made up as `Row Field 1`, …. */
const fieldHeaders = (fields: Grid, has: boolean, word: string): V[] => Array.from({ length: fields.cols }, (_, c) => (has ? gridAt(fields, c) : `${word} ${c + 1}`))

/** The depth of totals: left out, the grand total alone (Excel for Mac, the set `groupby`); past the fields' levels #VALUE!. */
function depthOf(n: number | undefined, levels: number, table: boolean): number | CErr {
  const d = n ?? 1
  if (Number.isNaN(d) || Math.abs(d) > levels || (table && Math.abs(d) > 1)) return E.VALUE
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
  const heads = headersOf(option(fieldHeadersMode), rowFields, values)
  if (isErr(heads)) return heads
  const start = heads.has ? 1 : 0
  const depth = depthOf(option(totalDepth), k, table)
  if (isErr(depth)) return depth
  const m = values.cols
  const measuredCols = m * fns.fns.length
  const sort = sortKeys(sortOrderGrid, k + measuredCols, k)
  if (isErr(sort)) return sort
  const rows = kept(filterGrid, n, start)
  if (isErr(rows)) return rows
  const agg = aggregator(cx, values, start, fns)
  return summary(() => {
    const axis = lines(rowFields, start, rows, k, { depth, table, sort }, (r, c) => agg(r, c, rows), cx)
    const out: V[][] = []
    const { headers, labels } = measuredLabels(values, heads.has, fns, !fns.down)
    const fieldsRow = heads.show ? fieldHeaders(rowFields, heads.has, 'Row Field') : new Array<V>(k).fill(EMPTY_CELL)
    if (fns.down) {
      // a column of functions: a line for each function, its label beside the fields
      if (heads.show) out.push([...fieldsRow, EMPTY_CELL, ...headers])
      for (const line of axis) fns.fns.forEach((_, f) => out.push([...line.keys, fns.labels[f], ...Array.from({ length: m }, (_, c) => agg(line.rows, f * m + c, rows))]))
      return rowsArray(out)
    }
    if (fns.fns.length > 1) out.push([...new Array<V>(k).fill(EMPTY_CELL), ...labels])
    if (heads.show) out.push([...fieldsRow, ...headers])
    for (const line of axis) out.push([...line.keys, ...Array.from({ length: measuredCols }, (_, c) => agg(line.rows, c, rows))])
    return rowsArray(out)
  })
}

/** The rows of `a` that are also in `b`. */
const meet = (a: readonly number[], b: ReadonlySet<number>) => a.filter((r) => b.has(r))

/**
 * PIVOTBY: the row fields' lines down, the column fields' lines across, each meeting's values aggregated (an empty cell
 * where no row meets); with headers shown, the column fields' headers above their keys and the row fields' beside the
 * values' own. With several value columns the total columns stay empty, as Excel for Mac leaves them (the set `groupby`).
 */
function pivotby(cx: Cx, rowFields: Grid, colFields: Grid, values: Grid, fn: Arg, fieldHeadersMode?: number, rowDepth?: number, rowSort?: Grid, colDepth?: number, colSort?: Grid, filterGrid?: Grid, relativeTo?: number): Arg {
  const fns = functionsOf(fn)
  if (isErr(fns)) return fns
  const n = rowFields.rows
  if (values.rows !== n || colFields.rows !== n) return E.VALUE
  const [k, j] = [rowFields.cols, colFields.cols]
  const heads = headersOf(option(fieldHeadersMode), rowFields, values)
  if (isErr(heads)) return heads
  const start = heads.has ? 1 : 0
  const rel = option(relativeTo) ?? 0
  if (!(rel >= 0 && rel <= 4)) return E.VALUE
  const [rd, cd] = [depthOf(option(rowDepth), k, false), depthOf(option(colDepth), j, false)]
  if (isErr(rd)) return rd
  if (isErr(cd)) return cd
  const m = values.cols
  const measuredCols = m * fns.fns.length
  const [rs, cs] = [sortKeys(rowSort, k + measuredCols, k), sortKeys(colSort, j + measuredCols, j)]
  if (isErr(rs)) return rs
  if (isErr(cs)) return cs
  const rows = kept(filterGrid, n, start)
  if (isErr(rows)) return rows
  const agg = aggregator(cx, values, start, fns)
  return summary(() => {
    const rowLines = lines(rowFields, start, rows, k, { depth: rd, table: false, sort: rs }, (r, c) => agg(r, c, rows), cx)
    const colLines = lines(colFields, start, rows, j, { depth: cd, table: false, sort: cs }, (r, c) => agg(r, c, rows), cx)
    const { headers, labels } = measuredLabels(values, heads.has, fns, !fns.down)
    // a column of functions repeats each row line for each function, its label beside the row fields
    const [per, side] = fns.down ? [m, 1] : [measuredCols, 0]
    const corner = new Array<V>(k + side).fill(EMPTY_CELL)
    const across = (cell: (c: Line, i: number) => V) => colLines.flatMap((c) => Array.from({ length: per }, (_, i) => cell(c, i)))
    const out: V[][] = []
    // the column fields' headers in one, as text joined by commas
    if (heads.show) {
      const names = fieldHeaders(colFields, heads.has, 'Column Field').map((v) => cx.rules.toText(v)).join(', ')
      out.push([...corner, ...across((c, i) => (c === colLines[0] && i === 0 ? names : EMPTY_CELL))])
    }
    for (let level = 0; level < j; level++) out.push([...corner, ...across((c) => c.keys[level])])
    if (!fns.down && fns.fns.length > 1) out.push([...corner, ...across((_, i) => labels[i])])
    if (heads.show) out.push([...fieldHeaders(rowFields, heads.has, 'Row Field'), ...new Array<V>(side).fill(EMPTY_CELL), ...across((_, i) => headers[i])])
    const sets = new Map<Line, Set<number>>()
    const setOf = (l: Line) => {
      let x = sets.get(l)
      if (!x) sets.set(l, (x = new Set(l.rows)))
      return x
    }
    for (const r of rowLines)
      for (let f = 0; f < (fns.down ? fns.fns.length : 1); f++) {
        const cells = across((c, i) => {
          if (m > 1 && c.kind !== 'item') return EMPTY_CELL
          const here = meet(r.rows, setOf(c))
          // the total a function of two parameters is relative to: the column's, the row's, the grand, a parent's
          const total = rel === 0 ? c.rows : rel === 1 ? r.rows : rel === 2 ? rows : rel === 3 ? meet(r.rows, new Set(c.parent)) : meet(c.rows, new Set(r.parent))
          return here.length ? agg(here, f * m + i, total) : EMPTY_CELL
        })
        out.push([...r.keys, ...(fns.down ? [fns.labels[f]] : []), ...cells])
      }
    return rowsArray(out)
  })
}

/** The optional arguments left out between commas take their defaults. */
const defaults = (from: number, to: number) => ({ at: Object.fromEntries(Array.from({ length: to - from + 1 }, (_, i) => [from + i, 'default' as const])) })

export const functions: Record<string, FnDef> = {
  GROUPBY: { args: 'G G R n? n? G? G? n?', omit: defaults(3, 7), cx: true, array: true, fn: groupby },
  PIVOTBY: { args: 'G G G R n? n? G? n? G? G? n?', omit: defaults(4, 10), cx: true, array: true, fn: pivotby },
}
