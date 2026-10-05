/**
 * Engine runtime: the value model, operators with type coercion, errors as values, ranges as slices of
 * sheet columns, the whitelisted functions.
 *
 * Only ROUND, EOMONTH, DATE and SUM come from formulajs: they pass the test against Excel documentation
 * and the oracle (DATE, EOMONTH and SUM through an adapter). SUMIF, SUMIFS, COUNTIFS (with a criteria
 * index), XLOOKUP, TEXT, IF, IFERROR and the volatile functions are our own: formulajs either lacks them
 * or gets them wrong. No platform imports: the code can be bundled and run outside Node.
 */
import * as fj from '@formulajs/formulajs'
import type { CellValue, ErrorCode } from './contract.js'
import { buildIndex, collator, matches, parseCrit, select, type Crit, type CritColumn, type CritIndex } from './criteria.js'
import { dateToSerial, MAX_ROWS } from './model.js'
import { formatText, generalText } from './text.js'

export class CErr {
  constructor(readonly code: ErrorCode) {}
}

export const E: Record<ErrorCode, CErr> = {
  DIV0: new CErr('DIV0'),
  NA: new CErr('NA'),
  NAME: new CErr('NAME'),
  NUM: new CErr('NUM'),
  REF: new CErr('REF'),
  VALUE: new CErr('VALUE'),
  CYCLE: new CErr('CYCLE'),
  SPILL: new CErr('SPILL'),
  ERROR: new CErr('ERROR'),
}

export type V = number | string | boolean | null | CErr

export interface SheetData {
  name: string
  /** Columns: cols[c][r]. */
  cols: V[][]
  /** Occupied rows (highest index + 1). */
  rows: number
}

/** A sheet rectangle; a whole column is cut to the occupied rows of the sheet when read. */
export class Rng {
  /** Key for the value and criteria-index caches. */
  readonly key: string
  constructor(
    readonly s: number,
    readonly r1: number,
    readonly c1: number,
    readonly r2: number,
    readonly c2: number,
  ) {
    this.key = `${s}:${r1}:${c1}:${r2}:${c2}`
  }
}

export const isErr = (x: unknown): x is CErr => x instanceof CErr

// ---------- formulajs ----------

const FJ_ERRORS: Record<string, ErrorCode> = { '#DIV/0!': 'DIV0', '#N/A': 'NA', '#NAME?': 'NAME', '#NUM!': 'NUM', '#REF!': 'REF', '#VALUE!': 'VALUE' }

/** A formulajs result in our model: Error → error, Date → serial number, NaN/∞ → #NUM!. */
function fromFj(x: unknown): V {
  if (x instanceof Error) return E[FJ_ERRORS[x.message] ?? 'ERROR']
  if (x instanceof Date) return localDateToSerial(x)
  if (typeof x === 'number') return Number.isFinite(x) ? x : E.NUM
  if (typeof x === 'string' || typeof x === 'boolean') return x
  return E.VALUE
}

/** formulajs reads and writes dates with local-time getters: serial numbers go through local midnight. */
function localDateToSerial(d: Date): number {
  const p = (n: number, w = 2) => String(n).padStart(w, '0')
  return dateToSerial(`${p(d.getFullYear(), 4)}-${p(d.getMonth() + 1)}-${p(d.getDate())}`)
}

function serialToLocalDate(serial: number): Date {
  const s = Math.floor(serial)
  const u = new Date(Date.UTC(1899, 11, 30) + (s < 61 ? s + 1 : s) * 86_400_000)
  return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate())
}

/** formulajs throws on some inputs: that is #VALUE!. */
function call(f: (...a: any[]) => unknown, ...a: unknown[]): V {
  try {
    return fromFj(f(...a))
  } catch {
    return E.VALUE
  }
}

// ---------- coercion ----------

export function toNum(v: V): number | CErr {
  if (typeof v === 'number') return v
  if (v === null) return 0
  if (typeof v === 'boolean') return v ? 1 : 0
  if (isErr(v)) return v
  const t = v.trim()
  if (t === '') return E.VALUE
  const n = Number(t)
  return Number.isFinite(n) ? n : E.VALUE
}

export function toText(v: V): string | CErr {
  if (typeof v === 'string') return v
  if (v === null) return ''
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (isErr(v)) return v
  return generalText(v)
}

export function toBool(v: V): boolean | CErr {
  if (typeof v === 'boolean') return v
  if (typeof v === 'number') return v !== 0
  if (v === null) return false
  if (isErr(v)) return v
  const u = v.toUpperCase()
  if (u === 'TRUE') return true
  if (u === 'FALSE') return false
  return E.VALUE
}

export function fromCellValue(v: CellValue): V {
  return v !== null && typeof v === 'object' ? E[v.error] : v
}

// ---------- context ----------

export interface RuntimeContext {
  sheets: SheetData[]
  /** The clock moment, a serial number with a day fraction (UTC). */
  now: number
  seed: number
  /** External function values by the canonical text of the call. */
  externals: Record<string, CellValue>
  /** Cache of large ranges (flat arrays) and criteria indexes; invalidated by edits. */
  flat: Map<string, V[]>
  index: Map<string, CritIndex>
}

export const BIG_RANGE = 256

/** Pseudo-random [0, 1) from the seed and the cell address: the same for the same clock. */
export function randAt(seed: number, sheet: string, row: number, col: number): number {
  let h = 0x811c9dc5 ^ (seed >>> 0)
  const key = `${sheet}\u0000${row}\u0000${col}`
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

  // ---------- operators ----------

  const arith = (a: V, b: V, f: (x: number, y: number) => number): V => {
    if (isErr(a)) return a
    if (isErr(b)) return b
    const x = toNum(a)
    if (isErr(x)) return x
    const y = toNum(b)
    if (isErr(y)) return y
    const z = f(x, y)
    return Number.isFinite(z) ? z : E.NUM
  }
  const add = (a: V, b: V) => arith(a, b, (x, y) => x + y)
  const sub = (a: V, b: V) => arith(a, b, (x, y) => x - y)
  const mul = (a: V, b: V) => arith(a, b, (x, y) => x * y)
  const div = (a: V, b: V): V => {
    if (isErr(a)) return a
    if (isErr(b)) return b
    const y = toNum(b)
    if (isErr(y)) return y
    if (y === 0) {
      const x = toNum(a)
      return isErr(x) ? x : E.DIV0
    }
    return arith(a, y, (x, z) => x / z)
  }
  const pow = (a: V, b: V) => arith(a, b, (x, y) => Math.pow(x, y))
  const neg = (a: V): V => {
    const x = toNum(a)
    return isErr(x) ? x : -x
  }
  const pct = (a: V): V => {
    const x = toNum(a)
    return isErr(x) ? x : x / 100
  }
  const cat = (a: V, b: V): V => {
    const x = toText(a)
    if (isErr(x)) return x
    const y = toText(b)
    return isErr(y) ? y : x + y
  }
  // Excel: numbers < text < booleans; text is case-insensitive; blank is 0, "" or FALSE
  const rank = (v: V) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2)
  const compare = (a: V, b: V): number => {
    if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0
    if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0
    const ra = rank(a)
    const rb = rank(b)
    if (ra !== rb) return ra - rb
    if (typeof a === 'string') return collator.compare(a, b as string)
    return (a as number) < (b as number) ? -1 : (a as number) > (b as number) ? 1 : 0
  }
  const cmp = (op: string, a: V, b: V): V => {
    if (isErr(a)) return a
    if (isErr(b)) return b
    const c = compare(a, b)
    return op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0
  }

  /** A value in a scalar position: the single cell of a range or #VALUE! (no spilling). */
  const scalar = (x: V | Rng): V => {
    if (!(x instanceof Rng)) return x
    const b = bounds(x)
    return b.h === 1 && b.w === 1 ? cell(x.s, b.r1, b.c1) : E.VALUE
  }
  const num = (x: V | Rng) => toNum(scalar(x))

  // ---------- criteria ----------

  const indexOf = (g: Rng): CritIndex => {
    let ix = ctx.index.get(g.key)
    if (!ix) ctx.index.set(g.key, (ix = buildIndex(flat(g))))
    return ix
  }

  const asRng = (x: V | Rng): Rng | CErr => (x instanceof Rng ? x : isErr(x) ? x : E.VALUE)

  /** Positions where all (range, criterion) pairs hold; #VALUE! when sizes differ. */
  const pick = (pairs: [V | Rng, V | Rng][], size?: { h: number; w: number }): { sel: number[]; size: { h: number; w: number } } | CErr => {
    const cols: CritColumn[] = []
    let sz = size
    for (const [rg, crit] of pairs) {
      const g = asRng(rg)
      if (isErr(g)) return g
      const b = bounds(g)
      sz ??= b
      if (b.h !== sz.h || b.w !== sz.w) return E.VALUE
      const v = scalar(crit)
      if (isErr(v)) return v
      cols.push({ flat: flat(g), crit: parseCrit(v), index: () => indexOf(g) })
    }
    return { sel: select(cols, sz!.h * sz!.w), size: sz! }
  }

  const sumAt = (g: Rng, sel: number[]): V => {
    const f = flat(g)
    let acc = 0
    for (const p of sel) {
      const v = f[p]
      if (typeof v === 'number') acc += v
      else if (isErr(v)) return v
    }
    return acc
  }

  const SUMIFS = (sum: V | Rng, ...rest: (V | Rng)[]): V => {
    const sg = asRng(sum)
    if (isErr(sg)) return sg
    const pairs: [V | Rng, V | Rng][] = []
    for (let i = 0; i < rest.length; i += 2) pairs.push([rest[i], rest[i + 1]])
    const r = pick(pairs, bounds(sg))
    return isErr(r) ? r : sumAt(sg, r.sel)
  }

  const COUNTIFS = (...args: (V | Rng)[]): V => {
    const pairs: [V | Rng, V | Rng][] = []
    for (let i = 0; i < args.length; i += 2) pairs.push([args[i], args[i + 1]])
    const r = pick(pairs)
    return isErr(r) ? r : r.sel.length
  }

  /** SUMIF: the sum range takes the size of the criteria range from its own top-left cell (as Excel does). */
  const SUMIF = (range: V | Rng, crit: V | Rng, sum?: V | Rng): V => {
    const cg = asRng(range)
    if (isErr(cg)) return cg
    const r = pick([[cg, crit]])
    if (isErr(r)) return r
    if (sum === undefined) return sumAt(cg, r.sel)
    const sg = asRng(sum)
    if (isErr(sg)) return sg
    const b = bounds(cg)
    const sized = new Rng(sg.s, sg.r1, sg.c1, sg.r1 + b.h - 1, sg.c1 + b.w - 1)
    return sumAt(sized, r.sel)
  }

  /** XLOOKUP: match modes 0, −1, 1, 2 (wildcard); search modes 1, −1, ±2 (binary search runs linearly, same answer on sorted data). */
  const XLOOKUP = (key: V | Rng, look: V | Rng, ret: V | Rng, notFound?: V | Rng, matchMode?: V | Rng, searchMode?: V | Rng): V => {
    const k = scalar(key)
    if (isErr(k)) return k
    const lg = asRng(look)
    if (isErr(lg)) return lg
    const rg = asRng(ret)
    if (isErr(rg)) return rg
    const lb = bounds(lg)
    const rb = bounds(rg)
    if (lb.w !== 1 && lb.h !== 1) return E.VALUE
    const vertical = lb.w === 1
    if ((vertical && rb.h !== lb.h) || (!vertical && rb.w !== lb.w)) return E.VALUE
    if ((vertical && rb.w !== 1) || (!vertical && rb.h !== 1)) return E.VALUE // an array result: no spilling
    const mm = matchMode === undefined ? 0 : num(matchMode)
    const sm = searchMode === undefined ? 1 : num(searchMode)
    if (isErr(mm)) return mm
    if (isErr(sm)) return sm
    if (![0, -1, 1, 2].includes(mm) || ![1, -1, 2, -2].includes(sm)) return E.VALUE
    const lf = flat(lg)
    const kk: V = k === null ? 0 : k
    const wild: Crit | null = mm === 2 && typeof kk === 'string' ? parseCrit('=' + kk) : null
    const order = sm > 0 ? lf.map((_, i) => i) : lf.map((_, i) => lf.length - 1 - i)
    let pos = -1
    let best: V = null
    for (const i of order) {
      const v = lf[i]
      if (v === null || isErr(v) || rank(v) !== rank(kk)) continue
      if (wild ? matches(v, wild) : compare(v, kk) === 0) {
        pos = i
        break
      }
      if (mm === -1 && compare(v, kk) < 0 && (pos < 0 || compare(v, best) > 0)) {
        pos = i
        best = v
      } else if (mm === 1 && compare(v, kk) > 0 && (pos < 0 || compare(v, best) < 0)) {
        pos = i
        best = v
      }
    }
    if (pos < 0) return notFound === undefined ? E.NA : scalar(notFound)
    return flat(rg)[pos] ?? null
  }

  const F: Record<string, (...a: any[]) => V> = {
    // in a range only numbers and errors count (text and booleans are skipped); a direct argument is coerced
    SUM: (...a: (V | Rng)[]) => {
      const args: unknown[] = []
      for (const x of a) {
        if (x instanceof Rng) {
          const f = flat(x)
          for (const v of f) if (isErr(v)) return v
          args.push(f.filter((v): v is number => typeof v === 'number'))
        } else {
          const n = toNum(x)
          if (isErr(n)) return n
          args.push(n)
        }
      }
      return call(fj.SUM, ...args)
    },
    SUMIF,
    SUMIFS,
    COUNTIFS,
    XLOOKUP,
    ROUND: (x: V | Rng, d: V | Rng) => {
      const n = num(x)
      if (isErr(n)) return n
      const k = num(d)
      if (isErr(k)) return k
      return call(fj.ROUND, n, Math.trunc(k))
    },
    EOMONTH: (x: V | Rng, m: V | Rng) => {
      const n = num(x)
      if (isErr(n)) return n
      const k = num(m)
      if (isErr(k)) return k
      if (n < 0) return E.NUM
      return call(fj.EOMONTH, serialToLocalDate(n), Math.trunc(k))
    },
    DATE: (y: V | Rng, m: V | Rng, d: V | Rng) => {
      const a = [y, m, d].map(num)
      for (const x of a) if (isErr(x)) return x
      const [yy, mo, dd] = (a as number[]).map(Math.trunc)
      // Excel: years 0–1899 are added to 1900; a result before 1900-01-01 is #NUM!
      const year = yy < 1900 ? yy + 1900 : yy
      if (year < 1900 || year > 9999) return E.NUM
      const v = call(fj.DATE, year, mo, dd)
      return typeof v === 'number' && v < 0 ? E.NUM : v
    },
    TEXT: (x: V | Rng, f: V | Rng) => {
      const v = scalar(x)
      if (isErr(v)) return v
      const t = toText(scalar(f))
      if (isErr(t)) return t
      return formatText(v ?? 0, t) ?? E.VALUE
    },
    TRUE: () => true,
    FALSE: () => false,
    TODAY: () => Math.floor(ctx.now),
    NOW: () => ctx.now,
  }

  const IF = (c: V | Rng, a: () => V | Rng, b: () => V | Rng): V => {
    const x = toBool(scalar(c))
    if (isErr(x)) return x
    return scalar(x ? a() : b())
  }
  const IFERR = (a: () => V | Rng, b: () => V | Rng): V => {
    const x = scalar(a())
    return isErr(x) ? scalar(b()) : x
  }
  const rand = (s: number, r: number, c: number) => randAt(ctx.seed, sheets[s].name, r, c)
  const randbetween = (s: number, r: number, c: number, lo: V | Rng, hi: V | Rng): V => {
    const a = num(lo)
    if (isErr(a)) return a
    const b = num(hi)
    if (isErr(b)) return b
    const x = Math.ceil(a)
    const y = Math.floor(b)
    if (y < x) return E.NUM
    return x + Math.floor(randAt(ctx.seed, sheets[s].name, r, c) * (y - x + 1))
  }
  const ext = (key: string): V => (key in ctx.externals ? fromCellValue(ctx.externals[key]) : E.NA)

  return { cell, flat, bounds, scalar, add, sub, mul, div, pow, neg, pct, cat, cmp, F, IF, IFERR, rand, randbetween, ext, E }
}

export type Runtime = ReturnType<typeof makeRuntime>
