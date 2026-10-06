/**
 * Database functions, this engine's own (formula.js reads the database by columns). As in Excel: records by
 * rows under a header row; criteria rows are OR, the conditions in a row AND; text without an operator
 * matches values that begin with it.
 */
import { gridAt, type FnDef, type Grid } from '../args.js'
import { errorCrit, matches, parseCrit, type Crit } from '../criteria.js'
import { E, isErr, toBool, toText, type CErr, type V } from '../values.js'

/** A criteria cell as a SUMIFS condition; bare text matches as a prefix (`App` selects `Apple`). */
function condition(v: V): Crit | null {
  if (v === null) return null
  if (isErr(v)) return errorCrit(v.code)
  if (typeof v !== 'string' || /^(<=|>=|<>|<|>|=)/.test(v) || v === '') return parseCrit(v)
  const c = parseCrit(v)
  return c.kind === 'str' || c.kind === 'wild' ? parseCrit(v + '*') : c
}

/** The column of a field name (any case) or number in the header row, or #VALUE!. */
function fieldIndex(db: Grid, field: V): number | CErr {
  if (isErr(field)) return field
  if (typeof field === 'number' || typeof field === 'boolean') {
    const i = Math.trunc(Number(field))
    return i >= 1 && i <= db.cols ? i - 1 : E.VALUE
  }
  return nameIndex(db, field)
}

/** The column whose header reads as `name` (any case), or #VALUE!; a criteria header is always a name, 45688 too. */
function nameIndex(db: Grid, name: V): number | CErr {
  if (isErr(name)) return name
  const text = (toText(name) as string).toLowerCase()
  for (let c = 0; c < db.cols; c++) {
    const h = gridAt(db, c)
    if (typeof h === 'string' ? h.toLowerCase() === text : toText(h) === text) return c
  }
  return E.VALUE
}

/** The field's values in the records selected by the criteria. */
function selected(db: Grid, field: V, criteria: Grid): V[] | CErr {
  const col = fieldIndex(db, field)
  if (isErr(col)) return col
  if (criteria.rows < 2) return E.VALUE
  // the database column each condition tests; under a header that is no field (−1) only TRUE holds, as a computed criterion
  const tests: { col: number; crit: Crit }[][] = []
  // past the occupied part every criteria row is blank: one of them selects every record
  for (let r = 1, end = Math.min(criteria.rows, criteria.h + 1); r < end; r++) {
    const row: { col: number; crit: Crit }[] = []
    for (let c = 0; c < criteria.cols; c++) {
      const v = gridAt(criteria, r * criteria.cols + c)
      const crit = condition(v)
      if (!crit) continue
      const at = nameIndex(db, gridAt(criteria, c))
      if (!isErr(at)) row.push({ col: at, crit })
      else if (toBool(v) !== true) row.push({ col: -1, crit })
    }
    tests.push(row)
  }
  const out: V[] = []
  for (let r = 1; r < db.h; r++) {
    const rec = (c: number) => gridAt(db, r * db.cols + c)
    if (tests.some((row) => row.every((t) => t.col >= 0 && matches(rec(t.col), t.crit)))) out.push(rec(col))
  }
  return out
}

const numbers = (vs: V[]) => vs.filter((v): v is number => typeof v === 'number')

/** Sum of squared deviations, in two passes. */
function squares(xs: number[]): number {
  const mean = xs.reduce((s, x) => s + x, 0) / xs.length
  return xs.reduce((s, x) => s + (x - mean) ** 2, 0)
}

/** A database function of the selected values; with `errors`, an error among them is the answer. */
const dfn = (f: (vs: V[]) => V, errors = true): FnDef => ({
  args: 'G a G',
  fn: (db: Grid, field: V, criteria: Grid) => {
    const vs = selected(db, field, criteria)
    if (isErr(vs)) return vs
    if (errors) {
      const e = vs.find(isErr)
      if (e) return e
    }
    return f(vs)
  },
})

const over = (f: (xs: number[]) => V) => dfn((vs) => f(numbers(vs)))

export const functions: Record<string, FnDef> = {
  DAVERAGE: over((xs) => (xs.length ? xs.reduce((s, x) => s + x, 0) / xs.length : E.DIV0)),
  // DCOUNT passes over errors, the other functions answer with them
  DCOUNT: dfn((vs) => numbers(vs).length, false),
  DCOUNTA: dfn((vs) => vs.filter((v) => v !== null).length, false),
  // one record, else #VALUE! (none or a blank field) or #NUM! (several)
  DGET: dfn((vs) => (vs.length > 1 ? E.NUM : vs.length === 0 || vs[0] === null ? E.VALUE : vs[0]), false),
  DMAX: over((xs) => xs.reduce((m, x) => (x > m ? x : m), xs.length ? -Infinity : 0)),
  DMIN: over((xs) => xs.reduce((m, x) => (x < m ? x : m), xs.length ? Infinity : 0)),
  DPRODUCT: over((xs) => (xs.length ? xs.reduce((p, x) => p * x, 1) : 0)),
  DSTDEV: over((xs) => (xs.length < 2 ? E.DIV0 : Math.sqrt(squares(xs) / (xs.length - 1)))),
  DSTDEVP: over((xs) => (xs.length < 1 ? E.DIV0 : Math.sqrt(squares(xs) / xs.length))),
  DSUM: over((xs) => xs.reduce((s, x) => s + x, 0)),
  DVAR: over((xs) => (xs.length < 2 ? E.DIV0 : squares(xs) / (xs.length - 1))),
  DVARP: over((xs) => (xs.length < 1 ? E.DIV0 : squares(xs) / xs.length)),
}
