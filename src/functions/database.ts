/**
 * Database functions, this engine's own (formula.js reads the database by columns). As in Excel: records by
 * rows under a header row; criteria rows are OR, the conditions in a row AND; text without an operator
 * matches values that begin with it.
 */
import { gridAt, type Cx, type FnDef, type Grid } from '../args.js'
import { errorCrit, matches, parseCrit, type Crit } from '../criteria.js'
import type { Rules } from '../locale/rules.js'
import { E, isErr, Rng, whole, type CErr, type V } from '../values.js'
import { max, min, product, sum } from './common.js'
import { mean, stdev, variance } from './statistical/common.js'

/** A criteria cell as a SUMIFS condition; bare text matches as a prefix (`App` selects `Apple`). */
function condition(v: V, rules: Rules): Crit | null {
  if (v === null) return null
  if (isErr(v)) return errorCrit(v.code)
  if (typeof v !== 'string' || /^(<=|>=|<>|<|>|=)/.test(v) || v === '') return parseCrit(v, rules)
  const c = parseCrit(v, rules)
  return c.kind === 'str' || c.kind === 'wild' ? parseCrit(v + '*', rules) : c
}

/** The column of a field name (any case) or number in the header row, or #VALUE!. */
function fieldIndex(db: Grid, field: V, rules: Rules): number | CErr {
  if (isErr(field)) return field
  if (typeof field === 'number' || typeof field === 'boolean') {
    const i = whole(Number(field))
    return i >= 1 && i <= db.cols ? i - 1 : E.VALUE
  }
  return nameIndex(db, field, rules)
}

/** The column whose header reads as `name` (any case), or #VALUE!; a criteria header is always a name, 45688 too. */
function nameIndex(db: Grid, name: V, rules: Rules): number | CErr {
  if (isErr(name)) return name
  const text = (rules.toText(name) as string).toLowerCase()
  for (let c = 0; c < db.cols; c++) {
    const h = gridAt(db, c)
    if (typeof h === 'string' ? h.toLowerCase() === text : rules.toText(h) === text) return c
  }
  return E.VALUE
}

/**
 * The field's values in the records selected by the criteria, each record as 0 without a field (DCOUNT and DCOUNTA
 * then count the records); a database or criteria without a row under the header is #VALUE!.
 */
function selected(db: Grid, field: V | undefined, criteria: Grid, rules: Rules): V[] | CErr {
  const col = field === undefined ? -1 : fieldIndex(db, field, rules)
  if (isErr(col)) return col
  if (db.rows < 2 || criteria.rows < 2) return E.VALUE
  // the database column each condition tests; under a header that is no field (−1) only TRUE holds, as a computed criterion
  const tests: { col: number; crit: Crit }[][] = []
  // past the occupied part every criteria row is blank: one of them selects every record
  for (let r = 1, end = Math.min(criteria.rows, criteria.h + 1); r < end; r++) {
    const row: { col: number; crit: Crit }[] = []
    for (let c = 0; c < criteria.cols; c++) {
      const v = gridAt(criteria, r * criteria.cols + c)
      const crit = condition(v, rules)
      if (!crit) continue
      const at = nameIndex(db, gridAt(criteria, c), rules)
      if (!isErr(at)) row.push({ col: at, crit })
      else if (rules.toBool(v) !== true) row.push({ col: -1, crit })
    }
    tests.push(row)
  }
  const out: V[] = []
  for (let r = 1; r < db.h; r++) {
    const rec = (c: number) => gridAt(db, r * db.cols + c)
    if (tests.some((row) => row.every((t) => t.col >= 0 && matches(rec(t.col), t.crit, rules)))) out.push(col < 0 ? 0 : rec(col))
  }
  return out
}

const numbers = (vs: V[]) => vs.filter((v): v is number => typeof v === 'number')

/**
 * The field as Excel reads it: a value as it is (an error is the answer), a reference to one cell by its value
 * (an error there is #VALUE!), a reference to several cells #VALUE!; an array is lifted.
 */
function fieldOf(cx: Cx, field: V | Rng | undefined): V | undefined {
  if (!(field instanceof Rng)) return field
  if (field.r2 > field.r1 || field.c2 > field.c1) return E.VALUE
  const v = cx.scalar(field)
  return isErr(v) ? E.VALUE : v
}

/** A database function of the selected values; with `errors`, an error among them is the answer. */
const dfn = (f: (vs: V[]) => V, errors = true): FnDef => ({
  args: 'G r G',
  cx: true,
  fn: (cx: Cx, db: Grid, field: V | Rng | undefined, criteria: Grid) => {
    const vs = selected(db, fieldOf(cx, field), criteria, cx.rules)
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
  DAVERAGE: over((xs) => (xs.length ? mean(xs) : E.DIV0)),
  // DCOUNT passes over errors, the other functions answer with them
  DCOUNT: dfn((vs) => numbers(vs).length, false),
  DCOUNTA: dfn((vs) => vs.filter((v) => v !== null).length, false),
  // one record, else #VALUE! (none or a blank field) or #NUM! (several)
  DGET: dfn((vs) => (vs.length > 1 ? E.NUM : vs.length === 0 || vs[0] === null ? E.VALUE : vs[0]), false),
  DMAX: over(max),
  DMIN: over(min),
  DPRODUCT: over(product),
  DSTDEV: over((xs) => stdev(xs, 1)),
  DSTDEVP: over((xs) => stdev(xs, 0)),
  DSUM: over(sum),
  DVAR: over((xs) => variance(xs, 1)),
  DVARP: over((xs) => variance(xs, 0)),
}
