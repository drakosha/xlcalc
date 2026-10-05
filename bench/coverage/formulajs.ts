/**
 * Вызов formulajs на случае замера двумя способами.
 *
 * Голая formulajs: аргументы как лежат в ячейках (одна ячейка значением, диапазон двумерным массивом,
 * пустая null, ошибка объектом ошибки formulajs, дата серийным номером), ответ как есть (Date
 * отмечается отдельно, с серийным номером по местному времени и долей суток).
 *
 * Через адаптер: общие правила рантайма движка (`src/runtime.ts`), обобщённые по видам аргументов:
 * - ошибка в скалярном аргументе и в числовом диапазоне (виды N, M, D) сразу ответ;
 * - число (`n`, `p`) через toNum, целое (`i`) ещё и с отбрасыванием дробной части (как ROUND),
 *   текст через toText, логическое через toBool;
 * - дата (`d`, `D`): серийный номер в местную полночь JS Date, время (`t`) с долей суток;
 *   отрицательный номер это #NUM! (как EOMONTH);
 * - свёртка чисел (`N`): из ссылки только числа, прямой аргумент приводится (правило SUM);
 * - остальные диапазоны двумерными массивами, ошибки внутри объектами ошибок formulajs;
 * - ответ: ошибка formulajs в код, Date в серийный номер её местной даты, NaN и бесконечность в
 *   #NUM!, исключение в #VALUE!, прочее (null, undefined, массив) в #VALUE!.
 */
import * as fj from '@formulajs/formulajs'
import { dateToSerial } from '../../src/model.js'
import { CErr, E, isErr, toBool, toNum, toText, type V } from '../../src/values.js'
import type { Case } from './cases.js'
import { kindAt, parseKinds } from './cases.js'
import { REGIONS, type Cell, type ErrCode } from './data.js'
import { SIGNATURES, type Arg } from './signatures.js'

/** Значение в выдаче замера. */
export type Val =
  | { t: 'n'; v: number }
  | { t: 's'; v: string }
  | { t: 'b'; v: boolean }
  | { t: 'e'; v: ErrCode }
  /** JS Date из голой formulajs: v серийный номер по местному времени. */
  | { t: 'date'; v: number }
  | { t: 'nan'; v: string }
  | { t: 'throw'; v: string }
  | { t: 'other'; v: string }

export interface FjResult {
  adapter: Val
  /** Адаптер получил исключение (ответ #VALUE!). */
  adapterThrew?: string
  bare: Val
}

const FJ_ERR: Record<string, ErrCode> = {
  '#NULL!': 'NULL',
  '#DIV/0!': 'DIV0',
  '#VALUE!': 'VALUE',
  '#REF!': 'REF',
  '#NAME?': 'NAME',
  '#NUM!': 'NUM',
  '#N/A': 'NA',
  '#GETTING_DATA': 'GETTING_DATA',
  '#CALC!': 'CALC',
}
const toFjErr: Record<string, Error> = {
  DIV0: fj.utils.errors.div0,
  NA: fj.utils.errors.na,
  VALUE: fj.utils.errors.value,
  NUM: fj.utils.errors.num,
  REF: fj.utils.errors.ref,
  NAME: fj.utils.errors.name,
  NULL: fj.utils.errors.nil,
}

const cellV = (c: Cell | null): V => {
  if (c === null) return null
  if (typeof c !== 'object') return c
  if ('error' in c) return E[c.error as keyof typeof E] ?? new CErr('ERROR')
  if ('date' in c) return c.date
  return c.text
}

interface Ref {
  rows: V[][]
  single: boolean
}

function resolve(a: Arg): V | Ref {
  if (typeof a === 'string' && a.startsWith('@')) {
    const r = REGIONS[a.slice(1)]
    const rows = r.rows.map((row) => row.map(cellV))
    return { rows, single: rows.length === 1 && rows[0].length === 1 }
  }
  return a
}

const isRef = (x: V | Ref): x is Ref => x !== null && typeof x === 'object' && 'rows' in x
const fjv = (v: V): unknown => (isErr(v) ? toFjErr[v.code] ?? fj.utils.errors.error : v)

// ---------- даты ----------

export function localDateToSerial(d: Date): number {
  const p = (n: number, w = 2) => String(n).padStart(w, '0')
  const day = dateToSerial(`${p(d.getFullYear(), 4)}-${p(d.getMonth() + 1)}-${p(d.getDate())}`)
  const ms = ((d.getHours() * 60 + d.getMinutes()) * 60 + d.getSeconds()) * 1000 + d.getMilliseconds()
  return day + ms / 86_400_000
}

export function serialToLocalDate(serial: number): Date {
  const s = Math.floor(serial)
  const u = new Date(Date.UTC(1899, 11, 30) + (s < 61 ? s + 1 : s) * 86_400_000)
  return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate(), 0, 0, 0, Math.round((serial - s) * 86_400_000))
}

// ---------- аргументы ----------

/** Голый вызов: одна ячейка значением, диапазон двумерным массивом, ошибки объектами formulajs. */
function bareArg(a: Arg): unknown {
  const x = resolve(a)
  if (!isRef(x)) return x
  return x.single ? fjv(x.rows[0][0]) : x.rows.map((r) => r.map(fjv))
}

const scalarOf = (x: V | Ref): V => (isRef(x) ? (x.single ? x.rows[0][0] : E.VALUE) : x)

/** Аргумент через адаптер по виду; CErr значит ответ этой ошибкой без вызова. */
function adaptArg(a: Arg, kind: string): unknown {
  const x = resolve(a)
  switch (kind) {
    case 'n':
    case 'p':
    case 'i': {
      const v = scalarOf(x)
      if (isErr(v)) return v
      const n = toNum(v)
      return isErr(n) ? n : kind === 'i' ? Math.trunc(n) : n
    }
    case 'd':
    case 't': {
      const v = scalarOf(x)
      if (isErr(v)) return v
      const n = toNum(v)
      if (isErr(n)) return n
      // дата: местная полночь (как EOMONTH в рантайме); время: с долей суток
      return n < 0 ? E.NUM : serialToLocalDate(kind === 'd' ? Math.floor(n) : n)
    }
    case 's': {
      const v = scalarOf(x)
      return isErr(v) ? v : toText(v)
    }
    case 'b': {
      const v = scalarOf(x)
      return isErr(v) ? v : toBool(v)
    }
    case 'N': {
      if (!isRef(x)) {
        if (isErr(x)) return x
        return toNum(x)
      }
      const flat = x.rows.flat()
      const bad = flat.find(isErr)
      if (bad) return bad
      return flat.filter((v): v is number => typeof v === 'number')
    }
    case 'M': {
      if (!isRef(x)) return x
      const bad = x.rows.flat().find(isErr)
      if (bad) return bad
      return x.rows.map((r) => [...r])
    }
    case 'D': {
      if (!isRef(x)) return typeof x === 'number' ? serialToLocalDate(x) : x
      const bad = x.rows.flat().find(isErr)
      if (bad) return bad
      return x.rows.map((r) => r.map((v) => (typeof v === 'number' ? serialToLocalDate(v) : v)))
    }
    case 'a':
    case 'c':
      if (!isRef(x)) return fjv(x)
      return x.single ? fjv(x.rows[0][0]) : x.rows.map((r) => r.map(fjv))
    default: // V, R
      if (!isRef(x)) return fjv(x)
      return x.rows.map((r) => r.map(fjv))
  }
}

// ---------- ответ ----------

/** Элемент или размер ответа-массива (одномерный массив считается колонкой). */
function applyPick(r: unknown, c: Case): unknown {
  if (c.pick === undefined || r instanceof Error) return r
  const m: unknown[][] | null = Array.isArray(r) ? (r.length && Array.isArray(r[0]) ? (r as unknown[][]) : (r as unknown[]).map((x) => [x])) : null
  const rows = m ? m.length : 1
  const cols = m ? Math.max(0, ...m.map((x) => x.length)) : 1
  if (c.pick === 'rows') return rows
  if (c.pick === 'cols') return cols
  const [i, j] = c.pick
  if (i > rows || j > cols) return fj.utils.errors.ref
  return m ? m[i - 1][j - 1] : r
}

function rawVal(x: unknown): Val {
  if (x instanceof Error) return { t: 'e', v: FJ_ERR[x.message] ?? 'ERROR' }
  if (x instanceof Date) return Number.isNaN(x.getTime()) ? { t: 'nan', v: 'Invalid Date' } : { t: 'date', v: localDateToSerial(x) }
  if (typeof x === 'number') return Number.isFinite(x) ? { t: 'n', v: x } : { t: 'nan', v: String(x) }
  if (typeof x === 'string') return { t: 's', v: x }
  if (typeof x === 'boolean') return { t: 'b', v: x }
  return { t: 'other', v: x === undefined ? 'undefined' : x === null ? 'null' : Array.isArray(x) ? 'array' : typeof x }
}

/** Ответ через адаптер (правила fromFj рантайма движка: дата в серийный номер местной полночи). */
function adapterVal(x: unknown): Val {
  const r = rawVal(x)
  if (r.t === 'date') return { t: 'n', v: Math.floor(r.v) }
  if (r.t === 'nan') return { t: 'e', v: 'NUM' }
  if (r.t === 'other') return { t: 'e', v: 'VALUE' }
  return r
}

const sigs = new Map(SIGNATURES.map((s) => [s.name, s]))
const getFn = (name: string) => name.split('.').reduce<any>((o, k) => o[k], fj) as (...a: unknown[]) => unknown
const msg = (e: unknown) => (e instanceof Error ? `${e.name}: ${e.message}` : String(e)).slice(0, 200)

export function evalCase(c: Case): FjResult {
  const fn = getFn(c.fn)
  const sig = sigs.get(c.fn)!
  const kinds = parseKinds(sig)

  let bare: Val
  try {
    bare = rawVal(applyPick(fn(...c.args.map(bareArg)), c))
  } catch (e) {
    bare = { t: 'throw', v: msg(e) }
  }

  const args = c.args.map((a, i) => adaptArg(a, kindAt(sig, kinds, i)))
  const early = args.find(isErr)
  if (early) return { adapter: { t: 'e', v: early.code as ErrCode }, bare }
  try {
    return { adapter: adapterVal(applyPick(fn(...args), c)), bare }
  } catch (e) {
    return { adapter: { t: 'e', v: 'VALUE' }, adapterThrew: msg(e), bare }
  }
}
