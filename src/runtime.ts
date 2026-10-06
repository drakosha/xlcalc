/**
 * Engine runtime: the evaluation context, operators with type coercion, ranges as slices of sheet columns,
 * criteria selection over cached indexes, and the whitelisted functions called through the argument layer
 * (`args.ts`). The functions themselves live in `functions/` by category. No platform imports: the code can
 * be bundled and run outside Node.
 */
import { makeCaller, type Arg, type Cx, type Grid } from './args.js'
import { liftCaller, liftLazy } from './arrays/call.js'
import { liftN, lifts } from './arrays/lift.js'
import { Arr } from './arrays/value.js'
import type { CellValue } from './contract.js'
import { buildIndex, errorCrit, matches, parseCrit, select, type CritColumn, type CritIndex } from './criteria.js'
import { DEFS, FUNCTIONS } from './functions.js'
import { power } from './functions/math.js'
import { MAX_TEXT } from './functions/text/common.js'
import { MAX_COLS, MAX_ROWS } from './model.js'
import { compareValues, E, fitNumber, fromCellValue, isErr, Rng, toNum, toText, type CErr, type V } from './values.js'

export interface SheetData {
  name: string
  /** Columns: cols[c][r]. */
  cols: V[][]
  /** Occupied rows (highest index + 1). */
  rows: number
}

export interface RuntimeContext {
  sheets: SheetData[]
  /** The clock moment, a serial number with a day fraction (UTC). */
  now: number
  seed: number
  /** External function values by the canonical text of the call. */
  externals: Record<string, CellValue>
  /** Cache of large ranges (flat arrays) and criteria indexes; invalidated by edits. */
  flat: Map<string, V[]>
  /** Cells whose formula calls SUBTOTAL or AGGREGATE, by sheet (row · MAX_COLS + column): those two skip them. */
  nested: Map<number, Set<number>>
  index: Map<string, CritIndex>
}

export const BIG_RANGE = 256

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
  const { sheets } = ctx

  const bounds = (g: Rng) => {
    const sh = sheets[g.s]
    const r2 = g.r2 === MAX_ROWS - 1 ? Math.max(sh.rows - 1, g.r1) : g.r2
    const c2 = Math.min(g.c2, Math.max(sh.cols.length - 1, g.c1))
    return { r1: g.r1, c1: g.c1, r2, c2, h: r2 - g.r1 + 1, w: c2 - g.c1 + 1 }
  }

  const cell = (s: number, r: number, c: number): V => sheets[s].cols[c]?.[r] ?? null

  /** Range values row by row; large ones are cached until an edit touches them. */
  const flat = (g: Rng): V[] => {
    const hit = ctx.flat.get(g.key)
    if (hit) return hit
    const b = bounds(g)
    const out: V[] = new Array(b.h * b.w)
    const cols = sheets[g.s].cols
    for (let c = 0; c < b.w; c++) {
      const col = cols[b.c1 + c]
      for (let r = 0; r < b.h; r++) out[r * b.w + c] = col?.[b.r1 + r] ?? null
    }
    if (out.length >= BIG_RANGE) ctx.flat.set(g.key, out)
    return out
  }

  /** Range values like `flat`, with the cells whose formula calls SUBTOTAL or AGGREGATE blank (not counted twice). */
  const flatOuter = (g: Rng): readonly V[] => {
    const marks = ctx.nested.get(g.s)
    if (!marks?.size) return flat(g)
    const b = bounds(g)
    let out: V[] | undefined
    for (const pos of marks) {
      const r = Math.floor(pos / MAX_COLS)
      const c = pos - r * MAX_COLS
      if (r < b.r1 || r > b.r2 || c < b.c1 || c > b.c2) continue
      out ??= flat(g).slice()
      out[(r - b.r1) * b.w + (c - b.c1)] = null
    }
    return out ?? flat(g)
  }

  /** A value in a scalar position: the single cell of a range or array, else #VALUE!. */
  const scalar = (x: Arg): V => {
    if (x instanceof Arr) return x.rows === 1 && x.cols === 1 ? x.at(0, 0) : E.VALUE
    if (!(x instanceof Rng)) return x
    return x.r1 === x.r2 && x.c1 === x.c2 ? cell(x.s, x.r1, x.c1) : E.VALUE
  }

  /** A reference as an array of its referenced size (the occupied part stored); values and arrays as they are. */
  const view = (x: Arg): V | Arr => {
    if (!(x instanceof Rng)) return x
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
    const c = compareValues(a, b)
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

  const indexOf = (g: Rng): CritIndex => {
    let ix = ctx.index.get(g.key)
    if (!ix) ctx.index.set(g.key, (ix = buildIndex(flat(g))))
    return ix
  }

  /** A range's values laid out row by row in `w` columns (its occupied part can be narrower). */
  const aligned = (g: Rng, w: number): readonly V[] => {
    const b = bounds(g)
    const f = flat(g)
    if (b.w === w) return f
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
      const c = isErr(v) ? errorCrit(v.code) : parseCrit(v)
      blank &&= matches(null, c)
      h = Math.max(h, bounds(rg).h)
      const f = aligned(rg, w)
      cols.push({ flat: f, crit: c, index: bounds(rg).w === w ? () => indexOf(rg) : () => buildIndex(f) })
    }
    const n = h * w
    return { sel: select(cols, n), rest: blank ? rows * w - n : 0, read: (g) => aligned(g, w) }
  }

  const cx: Cx = {
    get now() {
      return ctx.now
    },
    rand: (s, r, c, k) => randAt(ctx.seed, sheets[s].name, r, c, k),
    flat,
    flatOuter,
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
    F[name] = makeCaller(def, cx)
    if (!def.lazy) FL[name] = liftCaller(FUNCTIONS[name], def.cell === true, F[name], cx)
  }
  /** A lazy function whose lifting arguments at `check` can be arrays. */
  const LZ = (name: string, check: number[], ...thunks: (() => Arg)[]): Arg => liftLazy(F[name], FUNCTIONS[name], check, thunks, cx)

  const ext = (key: string): V => (key in ctx.externals ? fromCellValue(ctx.externals[key]) : E.NA)

  return { cell, flat, bounds, scalar, view, add, sub, mul, div, pow, neg, pct, cat, cmp, L1, L2, C, F, FL, LZ, ext, E, cx }
}

export type Runtime = ReturnType<typeof makeRuntime>
