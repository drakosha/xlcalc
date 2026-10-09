/**
 * Engine runtime: the evaluation context, operators with type coercion, ranges as slices of sheet columns,
 * criteria selection over cached indexes, and the whitelisted functions called through the argument layer
 * (`args.ts`). The functions themselves live in `functions/` by category. No platform imports: the code can
 * be bundled and run outside Node.
 */
import { makeCaller, type Arg, type Bounds, type Cx, type Grid, type HiddenRule } from './args.js'
import { liftCaller, liftLazy } from './arrays/call.js'
import { liftN, lifts } from './arrays/lift.js'
import { Arr } from './arrays/value.js'
import type { CellValue } from './contract.js'
import type { RefOp } from './syntax/ast.js'
import { buildIndex, errorCrit, matches, parseCrit, select, type CritColumn, type CritIndex } from './criteria.js'
import { allot, spend } from './exec/budget.js'
import { DEFS, FUNCTIONS } from './functions.js'
import { power } from './functions/math.js'
import { MAX_TEXT } from './functions/text/common.js'
import type { HiddenRows } from './layout/hidden.js'
import type { Rules } from './locale/rules.js'
import type { CacheBudget, RangeCache } from './range-cache.js'
import type { AggregateName, RangeAggregates } from './range-aggregate.js'
import { readsClock } from './locale/rules/data.js'
import { MAX_COLS, MAX_ROWS } from './model.js'
import { fillAt, fillBounds, fillFlat, type Fill } from './spill/fill.js'
import { Areas, E, fitNumber, fromCellValue, isErr, Rng, type CErr, type V } from './values.js'

export interface SheetData {
  name: string
  /** Columns: cols[c][r]. */
  cols: V[][]
  /** Occupied rows (highest index + 1). */
  rows: number
  /** The unstored cells of large spilled arrays (`spill/fill.ts`). */
  fills: Fill[]
}

export interface RuntimeContext {
  sheets: SheetData[]
  /** The clock moment, a serial number with a day fraction (UTC). */
  now: number
  seed: number
  /** The rules of computation of the book's locale (operators, criteria, function arguments). */
  rules: Rules
  /** External function values by the canonical text of the call. */
  externals: Record<string, CellValue>
  /** What `flat` and `index` hold together with other books' caches; trimmed between formulas. */
  cache: CacheBudget
  /** Cache of large ranges (flat arrays) and criteria indexes, from their second read; invalidated by edits. */
  flat: RangeCache<V[]>
  /** Cells whose formula calls SUBTOTAL or AGGREGATE, by sheet (row · MAX_COLS + column): those two skip them. */
  nested: Map<number, Set<number>>
  /** Hidden rows by sheet, which SUBTOTAL and AGGREGATE skip by their code. */
  hidden: Map<number, HiddenRows>
  index: RangeCache<CritIndex>
  aggregates: RangeAggregates
  /** A spill reference `E2#` (`docs/spill.md`): the range a cell's array spilled into, the cell alone, or #REF!. */
  spillRef: (s: number, r: number, c: number) => Rng | CErr
}

export const BIG_RANGE = 256

// the bytes a cache budget counts, as measured on Node (`docs/cache.md`): a value's slot in a range's array; a
// position in a criteria index, and a distinct value's key with its list
const VALUE_BYTES = 8
const POSITION_BYTES = 16
const KEY_BYTES = 80

/** Pseudo-random [0, 1) from the seed, the cell address and an array element (RANDARRAY): the same for the same clock. */
export function randAt(seed: number, sheet: string, row: number, col: number, k?: number): number {
  let h = 0x811c9dc5 ^ (seed >>> 0)
  const key = k === undefined ? `${sheet}\u0000${row}\u0000${col}` : `${sheet}\u0000${row}\u0000${col}\u0000${k}`
  for (let i = 0; i < key.length; i++) {
    h ^= key.charCodeAt(i)
    h = Math.imul(h, 0x01000193)
  }
  h = (h + 0x9e3779b9) | 0
  h = Math.imul(h ^ (h >>> 16), 0x85ebca6b)
  h = Math.imul(h ^ (h >>> 13), 0xc2b2ae35)
  h ^= h >>> 16
  return (h >>> 0) / 4_294_967_296
}

export function makeRuntime(ctx: RuntimeContext) {
  const { sheets, rules } = ctx
  const { toNum, toText, compareOp } = rules
  const shapes = new WeakMap<readonly V[], Bounds>()

  const bounds = (g: Rng) => {
    // A write outside the range can grow the sheet without changing this snapshot's stored shape.
    const hit = ctx.flat.get(g.key)
    const shape = hit && shapes.get(hit)
    if (shape) return shape
    const sh = sheets[g.s]
    const [r2, c2] = fillBounds(sh.fills, g, g.r2 === MAX_ROWS - 1 ? Math.max(sh.rows - 1, g.r1) : g.r2, Math.min(g.c2, Math.max(sh.cols.length - 1, g.c1)))
    return { r1: g.r1, c1: g.c1, r2, c2, h: r2 - g.r1 + 1, w: c2 - g.c1 + 1 }
  }

  const cell = (s: number, r: number, c: number): V => sheets[s].cols[c]?.[r] ?? fillAt(sheets[s].fills, r, c)

  /** Range values row by row; large ones read again are cached until an edit touches them or the budget drops them. */
  const flat = (g: Rng): V[] => {
    const hit = ctx.flat.get(g.key)
    if (hit) return hit
    const b = bounds(g)
    allot(b.h * b.w)
    spend(b.h * b.w)
    const out: V[] = new Array(b.h * b.w)
    const sh = sheets[g.s]
    for (let c = 0; c < b.w; c++) {
      const col = sh.cols[b.c1 + c]
      for (let r = 0; r < b.h; r++) out[r * b.w + c] = col?.[b.r1 + r] ?? null
    }
    if (sh.fills.length) fillFlat(sh.fills, b, out)
    if (out.length >= BIG_RANGE && ctx.flat.admits(g.key)) {
      shapes.set(out, b)
      ctx.flat.set(g.key, out, out.length * VALUE_BYTES)
    }
    return out
  }

  /** Range values like `flat`, blank where SUBTOTAL and AGGREGATE skip them (`Cx.counted`). */
  const counted = (g: Rng, nested: boolean, hidden: HiddenRule): readonly V[] => {
    const marks = nested ? ctx.nested.get(g.s) : undefined
    const rows = ctx.hidden.get(g.s)
    const skipRows = rows && (hidden === 'all' || (hidden === 'filtered' && rows.filtered))
    if (!marks?.size && !skipRows) return flat(g)
    const b = bounds(g)
    let out: V[] | undefined
    const blank = (r: number, c: number) => {
      out ??= flat(g).slice()
      out[(r - b.r1) * b.w + (c - b.c1)] = null
    }
    for (const pos of marks ?? []) {
      const r = Math.floor(pos / MAX_COLS)
      const c = pos - r * MAX_COLS
      if (r >= b.r1 && r <= b.r2 && c >= b.c1 && c <= b.c2) blank(r, c)
    }
    if (skipRows)
      for (const [lo, hi] of rows.spans) for (let r = Math.max(lo, b.r1); r <= Math.min(hi, b.r2); r++) for (let c = b.c1; c <= b.c2; c++) blank(r, c)
    return out ?? flat(g)
  }

  /** A value in a scalar position: the single cell of a range or array, else #VALUE!. */
  const scalar = (x: Arg): V => {
    if (x instanceof Rng) return x.r1 === x.r2 && x.c1 === x.c2 ? cell(x.s, x.r1, x.c1) : E.VALUE
    if (x instanceof Arr) return x.rows === 1 && x.cols === 1 ? x.at(0, 0) : E.VALUE
    return x instanceof Areas ? E.VALUE : x
  }

  /** A reference as an array of its referenced size (the occupied part stored); values and arrays as they are; several areas #VALUE!. */
  const view = (x: Arg): V | Arr => {
    if (!(x instanceof Rng)) return x instanceof Areas ? E.VALUE : x
    const b = bounds(x)
    return new Arr(x.r2 - x.r1 + 1, x.c2 - x.c1 + 1, b.h, b.w, flat(x))
  }

  const grid = (x: Arg): Grid => {
    if (x instanceof Arr) return x.fill === null ? x : x.full()
    const a = view(x)
    return a instanceof Arr ? a : { rows: 1, cols: 1, h: 1, w: 1, v: [a] }
  }

  // ---------- operators ----------

  // the left operand is converted before the right one counts: `"abc"-NA()` is #VALUE!, as in Excel
  const arith = (a: V, b: V, f: (x: number, y: number) => number | CErr): V => {
    const x = toNum(a)
    if (isErr(x)) return x
    const y = toNum(b)
    if (isErr(y)) return y
    const z = f(x, y)
    return isErr(z) ? z : fitNumber(z)
  }
  const add = (a: V, b: V) => arith(a, b, (x, y) => x + y)
  const sub = (a: V, b: V) => arith(a, b, (x, y) => x - y)
  const mul = (a: V, b: V) => arith(a, b, (x, y) => x * y)
  const div = (a: V, b: V) => arith(a, b, (x, y) => (y === 0 ? E.DIV0 : x / y))
  const pow = (a: V, b: V) => arith(a, b, power)
  const neg = (a: V): V => {
    const x = toNum(a)
    return isErr(x) ? x : -x
  }
  const pct = (a: V): V => {
    const x = toNum(a)
    return isErr(x) ? x : x / 100
  }
  // Excel cuts the joined text at its limit of 32,767 characters
  const cat = (a: V, b: V): V => {
    const x = toText(a)
    if (isErr(x)) return x
    const y = toText(b)
    if (isErr(y)) return y
    const s = x + y
    return s.length > MAX_TEXT ? s.slice(0, MAX_TEXT) : s
  }
  const cmp = (op: string, a: V, b: V): V => {
    if (isErr(a)) return a
    if (isErr(b)) return b
    const c = compareOp(a, b)
    return op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0
  }

  // ---------- operators over arrays (the code generator uses them where an operand can be an array) ----------

  // a one-cell reference is its value; any array, a 1 × 1 one too, lifts the operator
  const one = (x: Arg): Arg => (x instanceof Rng && !lifts(x) ? scalar(x) : x)
  const L1 = (f: (a: V) => V, x: Arg): Arg => {
    const a = one(x)
    return lifts(a) ? liftN([view(a)], ([v]) => f(v)) : f(a as V)
  }
  const L2 = (f: (a: V, b: V) => V, x: Arg, y: Arg): Arg => {
    const a = one(x)
    const b = one(y)
    return lifts(a) || lifts(b) ? liftN([view(a), view(b)], ([u, v]) => f(u, v)) : f(a as V, b as V)
  }
  /** Comparisons by operator, for L2. */
  const C: Record<string, (a: V, b: V) => V> = Object.fromEntries(['=', '<>', '<', '>', '<=', '>='].map((op) => [op, (a: V, b: V) => cmp(op, a, b)]))

  // ---------- criteria ----------

  // an index that read the clock's year (text dates without one) is built again by each formula, so each reads it
  const indexOf = (g: Rng, values: readonly V[]): CritIndex => {
    const hit = ctx.index.get(g.key)
    if (hit) return hit
    const [ix, clocked] = readsClock(rules.clock, () => buildIndex(values, rules))
    if (!clocked && ctx.index.admits(g.key)) {
      let positions = 0
      for (const list of ix.values()) positions += list.length
      ctx.index.set(g.key, ix, positions * POSITION_BYTES + ix.size * KEY_BYTES)
    }
    return ix
  }

  /** A range's values laid out row by row in `w` columns (its occupied part can be narrower). */
  const aligned = (g: Rng, w: number): readonly V[] => {
    const b = bounds(g)
    const f = flat(g)
    if (b.w === w) return f
    allot(b.h * w)
    const out: V[] = new Array(b.h * w).fill(null)
    for (let r = 0; r < b.h; r++) for (let c = 0; c < b.w; c++) out[r * w + c] = f[r * b.w + c]
    return out
  }

  const select_: Cx['select'] = (pairs, sum) => {
    const ref = sum ?? pairs[0]?.[0]
    if (!(ref instanceof Rng)) return isErr(ref) ? ref : E.VALUE
    const rows = ref.r2 - ref.r1 + 1
    const w = ref.c2 - ref.c1 + 1
    // the part any range occupies; past it every range is blank
    let h = sum ? bounds(sum).h : 0
    const cols: CritColumn[] = []
    let blank = true
    for (const [rg, crit] of pairs) {
      if (!(rg instanceof Rng)) return isErr(rg) ? rg : E.VALUE
      if (rg.r2 - rg.r1 + 1 !== rows || rg.c2 - rg.c1 + 1 !== w) return E.VALUE
      const v = scalar(crit)
      const c = isErr(v) ? errorCrit(v.code) : parseCrit(v, rules)
      blank &&= matches(null, c, rules)
      h = Math.max(h, bounds(rg).h)
      const f = aligned(rg, w)
      cols.push({ flat: f, crit: c, index: bounds(rg).w === w ? () => indexOf(rg, f) : () => buildIndex(f, rules) })
    }
    const n = h * w
    return { sel: select(cols, n, rules), rest: blank ? rows * w - n : 0, read: (g) => aligned(g, w) }
  }

  const cx: Cx = {
    get now() {
      return ctx.now
    },
    rules,
    rand: (s, r, c, k) => randAt(ctx.seed, sheets[s].name, r, c, k),
    flat,
    counted,
    bounds,
    scalar,
    grid,
    view,
    select: select_,
  }

  const F: Record<string, (...a: any[]) => Arg> = {}
  /** The callers that lift a call over arrays in one-value arguments (non-lazy functions). */
  const FL: Record<string, (...a: any[]) => Arg> = {}
  for (const [name, def] of Object.entries(DEFS)) {
    const caller = makeCaller(def, cx)
    // one range is read cell by cell, never made whole, so `limits.cells` does not bound it
    F[name] = ['SUM', 'COUNT', 'COUNTA', 'MIN', 'MAX', 'AVERAGE'].includes(name)
      ? (...args: Arg[]) => args.length === 1 && args[0] instanceof Rng
        ? ctx.aggregates.read(name as AggregateName, args[0], () => bounds(args[0] as Rng), cell) ?? caller(...args)
        : caller(...args)
      : name === 'COUNTIF' || name === 'COUNTIFS'
        ? (...args: Arg[]) => args.length === 2 && args[0] instanceof Rng
          ? ctx.aggregates.countIf(args[0], scalar(args[1]), rules, () => bounds(args[0] as Rng), cell)
          : caller(...args)
        : caller
    if (!def.lazy) FL[name] = liftCaller(FUNCTIONS[name], def.cell === true, F[name], cx)
  }
  /** A lazy function whose lifting arguments at `check` can be arrays. */
  const LZ = (name: string, check: number[], ...thunks: (() => Arg)[]): Arg => liftLazy(F[name], FUNCTIONS[name], check, thunks, cx)

  const ext = (key: string): V => (key in ctx.externals ? fromCellValue(ctx.externals[key]) : E.NA)

  const spill = (s: number, r: number, c: number) => ctx.spillRef(s, r, c)

  /**
   * `@x` at row `r`, column `c`: a reference's cell in that row or column (one row or column is its own), an array's
   * top-left value; a union is #VALUE!, a reference across sheets #REF!.
   */
  const I = (x: Arg, r: number, c: number): Arg => {
    if (x instanceof Arr) return x.at(0, 0)
    if (x instanceof Areas) return x.cube ? E.REF : E.VALUE
    if (!(x instanceof Rng)) return x
    const row = x.r1 === x.r2 ? x.r1 : r >= x.r1 && r <= x.r2 ? r : -1
    const col = x.c1 === x.c2 ? x.c1 : c >= x.c1 && c <= x.c2 ? c : -1
    return row < 0 || col < 0 ? E.VALUE : new Rng(x.s, row, col, row, col)
  }

  /** A union: references on one sheet as several areas, else #VALUE! (an error among them is the answer). */
  const U = (...xs: Arg[]): Arg => {
    const list: Rng[] = []
    for (const x of xs) {
      if (isErr(x)) return x
      if (x instanceof Rng) list.push(x)
      else if (x instanceof Areas && !x.cube) list.push(...x.list)
      else return E.VALUE
    }
    return list.every((g) => g.s === list[0].s) ? new Areas(list) : E.VALUE
  }

  /**
   * A reference operator: the intersection (`' '`) of the areas of two references, #NULL! where none meet, or the
   * range operator (`':'`), the rectangle around both. Anything but references on one sheet is #VALUE!.
   */
  const RO = (op: RefOp, x: Arg, y: Arg): Arg => {
    if (isErr(x)) return x
    if (isErr(y)) return y
    const parts = (z: Arg) => (z instanceof Rng ? [z] : z instanceof Areas && !z.cube ? z.list : null)
    const a = parts(x)
    const b = parts(y)
    if (!a || !b || [...a, ...b].some((g) => g.s !== a[0].s)) return E.VALUE
    if (op === ':') {
      const all = [...a, ...b]
      return new Rng(a[0].s, Math.min(...all.map((g) => g.r1)), Math.min(...all.map((g) => g.c1)), Math.max(...all.map((g) => g.r2)), Math.max(...all.map((g) => g.c2)))
    }
    const met: Rng[] = []
    for (const g of a)
      for (const h of b) {
        const [r1, c1, r2, c2] = [Math.max(g.r1, h.r1), Math.max(g.c1, h.c1), Math.min(g.r2, h.r2), Math.min(g.c2, h.c2)]
        if (r1 <= r2 && c1 <= c2) met.push(new Rng(g.s, r1, c1, r2, c2))
      }
    return !met.length ? E.NULL : met.length === 1 ? met[0] : new Areas(met)
  }

  return { cell, spill, flat, bounds, scalar, view, add, sub, mul, div, pow, neg, pct, cat, cmp, L1, L2, C, F, FL, LZ, ext, I, U, RO, E, cx }
}

export type Runtime = ReturnType<typeof makeRuntime>
