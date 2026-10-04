/**
 * Рантайм варианта C: модель значений, операторы с приведением типов, ошибки как значения, диапазоны
 * как срезы колонок листа, функции. Функции formulajs идут через адаптер «диапазон в двумерный
 * массив»; своё только то, чего в formulajs нет или что там неверно и нужно замеру (XLOOKUP,
 * SUMIFS/COUNTIFS с индексом), плюс волатильные от момента и зерна.
 *
 * Без импорта node:*: этот же код собирается в бандл для замера в QuickJS.
 */
import * as fj from '@formulajs/formulajs'
import { dateToSerial, MAX_ROWS, type ErrorCode } from './model'
import { randAt, randBetween } from './functions'
import { buildIndex, collator, matches, parseCrit, select, type Crit, type CritColumn, type CritIndex } from './criteria'

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
  /** Колонки: cols[c][r]. */
  cols: V[][]
  /** Занятых строк (максимальный индекс + 1). */
  rows: number
}

/** Ссылка на прямоугольник листа; целая колонка обрезается по занятым строкам листа при чтении. */
export class Rng {
  /** Ключ для кэшей: по нему же строятся индексы критериев. */
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

const FJ_ERRORS: Record<string, ErrorCode> = {
  '#DIV/0!': 'DIV0',
  '#N/A': 'NA',
  '#NAME?': 'NAME',
  '#NUM!': 'NUM',
  '#REF!': 'REF',
  '#VALUE!': 'VALUE',
  '#NULL!': 'ERROR',
  '#ERROR!': 'ERROR',
}

const isErr = (x: unknown): x is CErr => x instanceof CErr

/** Ответ formulajs в нашу модель: Error → CErr, Date → серийный номер, NaN/∞ → #NUM!. */
function fromFj(x: unknown): V {
  if (x instanceof Error) return E[FJ_ERRORS[x.message] ?? 'ERROR']
  if (x instanceof Date) return localDateToSerial(x)
  if (typeof x === 'number') return Number.isFinite(x) ? x : E.NUM
  if (x === undefined) return null
  if (Array.isArray(x)) return E.VALUE
  return x as V
}

/** Наше значение для formulajs: CErr → Error с текстом Excel. */
const ERROR_MSG: Record<ErrorCode, string> = { DIV0: '#DIV/0!', NA: '#N/A', NAME: '#NAME?', NUM: '#NUM!', REF: '#REF!', VALUE: '#VALUE!', CYCLE: '#ERROR!', SPILL: '#ERROR!', ERROR: '#ERROR!' }
function toFj(v: V): unknown {
  if (isErr(v)) {
    const e = new Error(ERROR_MSG[v.code])
    e.name = ERROR_MSG[v.code]
    return e
  }
  return v
}

function localDateToSerial(d: Date): number {
  const iso = `${String(d.getFullYear()).padStart(4, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
  return dateToSerial(iso)
}

/** Серийный номер в локальную полночь (formulajs читает даты локальными геттерами). */
function serialToLocalDate(serial: number): Date {
  const s = Math.floor(serial)
  const ms = Date.UTC(1899, 11, 30) + (s < 61 ? s + 1 : s) * 86_400_000
  const u = new Date(ms)
  return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate())
}

// ---------- приведение ----------

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
  return numText(v)
}

/** Число как текст в общем формате: до 15 значащих цифр. */
export function numText(n: number): string {
  if (n === 0) return '0'
  return String(Number(n.toPrecision(15)))
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

// ---------- контекст расчёта ----------

export interface RuntimeContext {
  sheets: SheetData[]
  now: number
  seed: number
  /** Режим SUMIFS/COUNTIFS: formulajs или своя с индексом. */
  ifs: 'fj' | 'idx'
  /** Кэш больших диапазонов (плоские массивы) и индексов критериев. */
  flat: Map<string, V[]>
  index: Map<string, CritIndex>
}

export const BIG_RANGE = 256

export function makeRuntime(ctx: RuntimeContext) {
  const { sheets } = ctx

  const bounds = (g: Rng) => {
    const sh = sheets[g.s]
    const r2 = g.r2 === MAX_ROWS - 1 ? Math.min(g.r2, Math.max(sh.rows - 1, g.r1)) : g.r2
    const c2 = Math.min(g.c2, Math.max(sh.cols.length - 1, g.c1))
    return { r1: g.r1, c1: g.c1, r2, c2, h: r2 - g.r1 + 1, w: c2 - g.c1 + 1 }
  }

  const cell = (s: number, r: number, c: number): V => sheets[s].cols[c]?.[r] ?? null

  /** Значения диапазона построчно; большие кэшируются до правки, которая их задевает. */
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

  /** Адаптер «диапазон в двумерный массив» для formulajs; `numbersOnly` убирает текст и логические. */
  const grid = (g: Rng, numbersOnly = false): unknown[][] => {
    const b = bounds(g)
    const f = flat(g)
    const out: unknown[][] = []
    for (let r = 0; r < b.h; r++) {
      const row: unknown[] = []
      for (let c = 0; c < b.w; c++) {
        const v = f[r * b.w + c]
        row.push(numbersOnly ? (typeof v === 'number' || isErr(v) ? toFj(v) : null) : toFj(v))
      }
      out.push(row)
    }
    return out
  }

  // ---------- операторы ----------

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
  // Excel: числа < текст < логические; текст без учёта регистра; пусто как 0 или ""
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
    switch (op) {
      case '=':
        return c === 0
      case '<>':
        return c !== 0
      case '<':
        return c < 0
      case '>':
        return c > 0
      case '<=':
        return c <= 0
      default:
        return c >= 0
    }
  }

  /** Диапазон в скалярном месте (без растекания в C v1): одна ячейка или #VALUE!. */
  const scalar = (x: V | Rng): V => {
    if (!(x instanceof Rng)) return x
    const b = bounds(x)
    return b.h === 1 && b.w === 1 ? cell(x.s, b.r1, b.c1) : E.VALUE
  }

  // ---------- критерии (общие с вариантом H со своей SUMIFS) ----------

  const indexOf = (g: Rng): CritIndex => {
    let ix = ctx.index.get(g.key)
    if (!ix) ctx.index.set(g.key, (ix = buildIndex(flat(g))))
    return ix
  }

  /** Отбор позиций по парам (диапазон, критерий); #VALUE! при разных размерах. */
  const pick = (pairs: [Rng, V][], size: { h: number; w: number }): number[] | CErr => {
    const cols: CritColumn[] = []
    for (const [g, v] of pairs) {
      const b = bounds(g)
      if (b.h !== size.h || b.w !== size.w) return E.VALUE
      if (isErr(v)) return v
      cols.push({ flat: flat(g), crit: parseCrit(v), index: () => indexOf(g) })
    }
    return select(cols, size.h * size.w)
  }

  const asRng = (x: V | Rng): Rng | CErr => (x instanceof Rng ? x : isErr(x) ? x : E.VALUE)

  const sumifsIdx = (sum: V | Rng, ...rest: (V | Rng)[]): V => {
    const sg = asRng(sum)
    if (isErr(sg)) return sg
    if (rest.length === 0 || rest.length % 2) return E.VALUE
    const pairs: [Rng, V][] = []
    for (let i = 0; i < rest.length; i += 2) {
      const g = asRng(rest[i])
      if (isErr(g)) return g
      pairs.push([g, scalar(rest[i + 1])])
    }
    const b = bounds(sg)
    const sel = pick(pairs, b)
    if (isErr(sel)) return sel
    const f = flat(sg)
    let acc = 0
    for (const p of sel) {
      const v = f[p]
      if (typeof v === 'number') acc += v
      else if (isErr(v)) return v
    }
    return acc
  }

  const countifsIdx = (...args: (V | Rng)[]): V => {
    if (args.length === 0 || args.length % 2) return E.VALUE
    const pairs: [Rng, V][] = []
    for (let i = 0; i < args.length; i += 2) {
      const g = asRng(args[i])
      if (isErr(g)) return g
      pairs.push([g, scalar(args[i + 1])])
    }
    const sel = pick(pairs, bounds(pairs[0][0]))
    return isErr(sel) ? sel : sel.length
  }

  // ---------- функции ----------

  /** Аргумент formulajs: диапазон двумерным массивом, скаляр как есть. */
  const arg = (x: V | Rng, numbersOnly = false): unknown => (x instanceof Rng ? grid(x, numbersOnly) : toFj(x))
  // formulajs на части входов бросает исключение (TEXT с форматом даты): это #VALUE!
  const call = (f: (...a: any[]) => unknown, ...a: unknown[]): V => {
    try {
      return fromFj(f(...a))
    } catch {
      return E.VALUE
    }
  }

  const xlookup = (key: V | Rng, look: V | Rng, ret: V | Rng, notFound?: V | Rng, matchMode?: V | Rng, searchMode?: V | Rng): V => {
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
    const mm = matchMode === undefined ? 0 : toNum(scalar(matchMode))
    const sm = searchMode === undefined ? 1 : toNum(scalar(searchMode))
    if (isErr(mm)) return mm
    if (isErr(sm)) return sm
    if (mm !== 0 && mm !== 2) return E.VALUE // прототип: точное совпадение и шаблон
    if (sm !== 1 && sm !== -1) return E.VALUE
    const lf = flat(lg)
    const crit: Crit = mm === 2 && typeof k === 'string' ? parseCrit('=' + k) : { op: '=', kind: 'num', n: 0 }
    const hit = (v: V) => (mm === 2 && typeof k === 'string' ? matches(v, crit) : v !== null && !isErr(v) && compare(v, k) === 0 && rank(v) === rank(k as V))
    let pos = -1
    if (sm === 1) {
      for (let i = 0; i < lf.length; i++)
        if (hit(lf[i])) {
          pos = i
          break
        }
    } else {
      for (let i = lf.length - 1; i >= 0; i--)
        if (hit(lf[i])) {
          pos = i
          break
        }
    }
    if (pos < 0) return notFound === undefined ? E.NA : scalar(notFound)
    const rf = flat(rg)
    if (rb.w === 1 || rb.h === 1) return rf[pos] ?? null
    return E.VALUE // возврат строки или колонки массивом: растекание, в C v1 нет
  }

  const F: Record<string, (...a: any[]) => V> = {
    SUM: (...a: (V | Rng)[]) => {
      // в диапазоне считаются только числа (текст и логические пропускаются), прямой аргумент приводится
      const args: unknown[] = []
      for (const x of a) {
        if (x instanceof Rng) args.push(grid(x, true))
        else {
          const n = toNum(x)
          if (isErr(n)) return n
          args.push(n)
        }
      }
      return call(fj.SUM, ...args)
    },
    SUMIFS: (...a: (V | Rng)[]) => (ctx.ifs === 'idx' ? sumifsIdx(a[0], ...a.slice(1)) : call(fj.SUMIFS, ...a.map((x) => arg(x)))),
    COUNTIFS: (...a: (V | Rng)[]) => (ctx.ifs === 'idx' ? countifsIdx(...a) : call(fj.COUNTIFS, ...a.map((x) => arg(x)))),
    XLOOKUP: xlookup,
    ROUND: (x: V | Rng, d: V | Rng) => {
      const n = toNum(scalar(x))
      if (isErr(n)) return n
      const k = toNum(scalar(d))
      if (isErr(k)) return k
      return call(fj.ROUND, n, k)
    },
    EOMONTH: (x: V | Rng, m: V | Rng) => {
      const n = toNum(scalar(x))
      if (isErr(n)) return n
      const k = toNum(scalar(m))
      if (isErr(k)) return k
      if (n < 0) return E.NUM
      return call(fj.EOMONTH, serialToLocalDate(n), Math.trunc(k))
    },
    DATE: (y: V | Rng, m: V | Rng, d: V | Rng) => {
      const a = [y, m, d].map((x) => toNum(scalar(x)))
      for (const x of a) if (isErr(x)) return x
      return call(fj.DATE, ...(a as number[]).map(Math.trunc))
    },
    TEXT: (x: V | Rng, f: V | Rng) => {
      const v = scalar(x)
      if (isErr(v)) return v
      const t = toText(scalar(f))
      if (isErr(t)) return t
      return call(fj.TEXT, v, t)
    },
    TRUE: () => true,
    FALSE: () => false,
    TODAY: () => Math.floor(ctx.now),
    NOW: () => ctx.now,
  }

  const rand = (s: number, r: number, c: number) => randAt(ctx.seed, sheets[s].name, r, c)
  const randbetween = (s: number, r: number, c: number, lo: V | Rng, hi: V | Rng): V => {
    const a = toNum(scalar(lo))
    if (isErr(a)) return a
    const b = toNum(scalar(hi))
    if (isErr(b)) return b
    if (b < a) return E.NUM
    return randBetween(ctx.seed, sheets[s].name, r, c, a, b)
  }

  return { cell, flat, scalar, add, sub, mul, div, pow, neg, pct, cat, cmp, toBool, isErr, F, rand, randbetween, Rng, E }
}

export type Runtime = ReturnType<typeof makeRuntime>
