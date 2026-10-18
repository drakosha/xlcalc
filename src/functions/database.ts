/**
 * Database functions, this engine's own (formula.js reads the database by columns). As in Excel: records by
 * rows under a header row; criteria rows are OR, the conditions in a row AND; text without an operator
 * matches values that begin with it. A database given as a reference selects each criteria row through the
 * runtime's selection, as SUMIFS does (kept selections, indexes); one given as an array is passed over record by record.
 */
import { gridAt, type Cx, type FnDef, type Grid } from '../args.js'
import { isArea } from '../arrays/lift.js'
import { errorCrit, matches, parseCrit, type Crit } from '../criteria.js'
import type { Rules } from '../locale/rules.js'
import { E, isErr, Rng, whole, type CErr, type V } from '../values.js'
import { max, min, product, sum } from './common.js'
import { meanOrDiv0, numbersOf, stdev, variance } from './statistical/common.js'

/** What tests count of the database functions (`watchRecords`): the records checked one by one. Nothing is counted unless a test asks. */
export interface RecordWatch {
  records: number
}
let watch: RecordWatch | undefined

/** Counts the records the database functions check one by one from now on into `w`; undefined stops. For tests. */
export function watchRecords(w?: RecordWatch): void {
  watch = w
}

/** A criteria cell as the value of a SUMIFS criterion: bare text that reads as text matches as a prefix (`App` selects `Apple`). */
function conditionValue<T extends V>(v: T, rules: Rules): T | string {
  if (typeof v !== 'string' || /^(<=|>=|<>|<|>|=)/.test(v) || v === '') return v
  const { kind } = parseCrit(v, rules)
  return kind === 'str' || kind === 'wild' ? v + '*' : v
}

/** A criteria cell as a SUMIFS condition. */
const condition = (v: V, rules: Rules): Crit => (isErr(v) ? errorCrit(v.code) : parseCrit(conditionValue(v, rules), rules))

/** The column of a field name (any case) or number in the header row, or #VALUE!. */
function fieldIndex(db: Grid, field: V, rules: Rules): number | CErr {
  if (isErr(field)) return field
  if (typeof field === 'number' || typeof field === 'boolean') {
    const i = whole(Number(field))
    return i >= 1 && i <= db.cols ? i - 1 : E.VALUE
  }
  return nameIndex(db, field, rules)
}

/**
 * The first column whose header reads as `name` in any case, or #VALUE!: a criteria header is always a name, 45688
 * too, and the text `true` names a header TRUE as TRUE names a header `true`.
 */
function nameIndex(db: Grid, name: V, rules: Rules): number | CErr {
  if (isErr(name)) return name
  const text = (rules.toText(name) as string).toLowerCase()
  for (let c = 0; c < db.cols; c++) {
    const h = gridAt(db, c)
    if ((typeof h === 'string' ? h : (rules.toText(h) as string)).toLowerCase() === text) return c
  }
  return E.VALUE
}

/** A condition of a criteria row: the database column it tests (−1 under a header that names no field) and the cell's value. */
interface Test {
  col: number
  v: V
}

/**
 * The criteria rows, each its conditions, the columns by the database's header row; under a header that is no field
 * only TRUE holds, as a computed criterion. Past the occupied part every criteria row is blank: one of them selects every record.
 */
function criteriaRows(header: Grid, criteria: Grid, rules: Rules): Test[][] {
  const rows: Test[][] = []
  for (let r = 1, end = Math.min(criteria.rows, criteria.h + 1); r < end; r++) {
    const row: Test[] = []
    for (let c = 0; c < criteria.cols; c++) {
      const v = gridAt(criteria, r * criteria.cols + c)
      if (v === null) continue
      const at = nameIndex(header, gridAt(criteria, c), rules)
      if (!isErr(at)) row.push({ col: at, v })
      else if (rules.toBool(v) !== true) row.push({ col: -1, v })
    }
    rows.push(row)
  }
  return rows
}

/**
 * The field's values in the records selected by the criteria, each record as 0 without a field (DCOUNT and DCOUNTA
 * then count the records); a database or criteria without a row under the header is #VALUE!. A database given as an array.
 */
function selected(db: Grid, field: V | undefined, criteria: Grid, rules: Rules): V[] | CErr {
  const col = field === undefined ? -1 : fieldIndex(db, field, rules)
  if (isErr(col)) return col
  if (db.rows < 2 || criteria.rows < 2) return E.VALUE
  const tests = criteriaRows(db, criteria, rules).map((row) => row.map((t) => ({ col: t.col, crit: condition(t.v, rules) })))
  const out: V[] = []
  if (watch) watch.records += Math.max(0, db.h - 1)
  for (let r = 1; r < db.h; r++) {
    const rec = (c: number) => gridAt(db, r * db.cols + c)
    if (tests.some((row) => row.every((t) => t.col >= 0 && matches(rec(t.col), t.crit, rules)))) out.push(col < 0 ? 0 : rec(col))
  }
  return out
}

/** The positions `from`…`to` − 1. */
const span = (from: number, to: number) => Array.from({ length: Math.max(0, to - from) }, (_, i) => from + i)

/** Positions in either of two ascending lists, ascending, each once. */
function union(a: readonly number[], b: readonly number[]): number[] {
  const out: number[] = []
  let i = 0
  let j = 0
  while (i < a.length || j < b.length) {
    const x = j === b.length || (i < a.length && a[i] <= b[j]) ? a[i++] : b[j++]
    if (out.length === 0 || out[out.length - 1] !== x) out.push(x)
  }
  return out
}

/**
 * `selected` for a database given as a reference: each criteria row selects through the runtime (`Cx.select`) over the
 * database's columns, the rows merged in order, so the records come as the pass over them gives them.
 */
function selectedIn(cx: Cx, db: Rng, field: V | undefined, criteria: Grid): V[] | CErr {
  const { rules } = cx
  const header = cx.grid(new Rng(db.s, db.r1, db.c1, db.r1, db.c2))
  const col = field === undefined ? -1 : fieldIndex(header, field, rules)
  if (isErr(col)) return col
  if (db.r2 === db.r1 || criteria.rows < 2) return E.VALUE
  // the records the pass would visit: the database's occupied part under its header
  const n = cx.bounds(db).h - 1
  const rows = db.r2 - db.r1
  const column = (c: number) => new Rng(db.s, db.r1 + 1, db.c1 + c, db.r2, db.c1 + c)
  let sel: readonly number[] | undefined
  for (const row of criteriaRows(header, criteria, rules)) {
    if (row.some((t) => t.col < 0)) continue
    if (!row.length) {
      sel = span(0, n)
      break
    }
    const s = cx.select(row.map((t) => [column(t.col), conditionValue(t.v, rules)]))
    if (isErr(s)) return s
    // the selection reads the part its columns occupy, which can pass the database's or stop short of it (blank there)
    const blank = rows - s.rest
    let got = s.sel.length && s.sel[s.sel.length - 1] >= n ? s.sel.filter((p) => p < n) : s.sel
    if (s.rest && blank < n) got = got.concat(span(blank, n))
    sel = sel ? union(sel, got) : got
  }
  if (!sel) return []
  if (col < 0) return new Array<V>(sel.length).fill(0)
  const values = cx.flat(column(col))
  return sel.map((p) => values[p] ?? null)
}

/**
 * The field as Excel reads it: a value as it is (an error is the answer), a reference to one cell by its value
 * (an error there is #VALUE!), a reference to several cells #VALUE!; an array is lifted.
 */
function fieldOf(cx: Cx, field: V | Rng | undefined): V | undefined {
  if (!(field instanceof Rng)) return field
  if (isArea(field)) return E.VALUE
  const v = cx.scalar(field)
  return isErr(v) ? E.VALUE : v
}

/** A database function of the selected values. */
const dfn = (f: (vs: V[]) => V): FnDef => ({
  args: 'D r G',
  cx: true,
  fn: (cx: Cx, db: Grid | Rng, field: V | Rng | undefined, criteria: Grid) => {
    const at = fieldOf(cx, field)
    const vs = db instanceof Rng ? selectedIn(cx, db, at, criteria) : selected(db, at, criteria, cx.rules)
    return isErr(vs) ? vs : f(vs)
  },
})

/** A database function of the selected numbers; an error among the values is the answer. */
const over = (f: (xs: number[]) => V) =>
  dfn((vs) => {
    const xs = numbersOf(vs)
    return isErr(xs) ? xs : f(xs)
  })

export const functions: Record<string, FnDef> = {
  DAVERAGE: over(meanOrDiv0),
  // DCOUNT passes over errors, the other functions answer with them
  DCOUNT: dfn((vs) => vs.filter((v) => typeof v === 'number').length),
  DCOUNTA: dfn((vs) => vs.filter((v) => v !== null).length),
  // one record, else #VALUE! (none or a blank field) or #NUM! (several)
  DGET: dfn((vs) => (vs.length > 1 ? E.NUM : vs.length === 0 || vs[0] === null ? E.VALUE : vs[0])),
  DMAX: over(max),
  DMIN: over(min),
  DPRODUCT: over(product),
  DSTDEV: over((xs) => stdev(xs, 1)),
  DSTDEVP: over((xs) => stdev(xs, 0)),
  DSUM: over(sum),
  DVAR: over((xs) => variance(xs, 1)),
  DVARP: over((xs) => variance(xs, 0)),
}
