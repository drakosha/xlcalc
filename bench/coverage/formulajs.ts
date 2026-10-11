/**
 * formulajs called on a measurement case in two ways.
 *
 * Bare formulajs: arguments as they lie in the cells (one cell as its value, a range as a two-dimensional
 * array, a blank as null, an error as a formulajs error object, a date as a serial number), the answer as it
 * is (a Date is marked apart, as a serial number in local time with the fraction of the day).
 *
 * Through the adapter: the engine's runtime rules (`src/runtime.ts`), generalised by argument kind:
 * - an error in a scalar argument or in a numeric range (kinds N, M, D) is the answer at once;
 * - a number (`n`, `p`) through toNum, a whole number (`i`) also truncated (as ROUND does), text through
 *   toText, a boolean through toBool;
 * - a date (`d`, `D`): the serial number as a JS Date at local midnight, a time (`t`) with the fraction of the
 *   day; a negative serial number is #NUM! (as in EOMONTH);
 * - a fold of numbers (`N`): only numbers from a reference, a direct argument converted (the SUM rule);
 * - other ranges as two-dimensional arrays, errors inside them as formulajs error objects;
 * - the answer: a formulajs error to its code, a Date to the serial number of its local date, NaN and
 *   infinity to #NUM!, an exception to #VALUE!, anything else (null, undefined, an array) to #VALUE!.
 */
import * as fj from '@formulajs/formulajs'
import { dateToSerial } from '../../src/model.js'
import { ERROR_CODE, ERROR_TEXT, TYPED_ERRORS } from '../../src/syntax/errors.js'
import { CErr, E, isErr, toBool, toNum, toText, type V } from '../../src/values.js'
import type { Case } from './cases.js'
import { kindAt, parseKinds } from './cases.js'
import { REGIONS, type Cell, type ErrCode } from './data.js'
import { SIGNATURES, type Arg } from './signatures.js'

/** A value in the measurement's output. */
export type Val =
  | { t: 'n'; v: number }
  | { t: 's'; v: string }
  | { t: 'b'; v: boolean }
  | { t: 'e'; v: ErrCode }
  /** A JS Date from bare formulajs: v is the serial number in local time. */
  | { t: 'date'; v: number }
  | { t: 'nan'; v: string }
  | { t: 'throw'; v: string }
  | { t: 'other'; v: string }

export interface FjResult {
  adapter: Val
  /** The adapter caught an exception (the answer is #VALUE!). */
  adapterThrew?: string
  bare: Val
}

/** formulajs's errors by their text: those typed into a cell, `#GETTING_DATA` and `#CALC!` (no `#SPILL!`). */
const FJ_ERR: Record<string, ErrCode> = { ...Object.fromEntries(TYPED_ERRORS.map((t) => [t, ERROR_CODE[t] as ErrCode])), '#GETTING_DATA': 'GETTING_DATA', [ERROR_TEXT.CALC]: 'CALC' }
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

// ---------- dates ----------

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

// ---------- arguments ----------

/** A bare call: one cell as its value, a range as a two-dimensional array, errors as formulajs objects. */
function bareArg(a: Arg): unknown {
  const x = resolve(a)
  if (!isRef(x)) return x
  return x.single ? fjv(x.rows[0][0]) : x.rows.map((r) => r.map(fjv))
}

const scalarOf = (x: V | Ref): V => (isRef(x) ? (x.single ? x.rows[0][0] : E.VALUE) : x)

/** An argument through the adapter by kind; a CErr means that error is the answer without a call. */
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
      // a date: local midnight (as EOMONTH in the runtime); a time: with the fraction of the day
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

// ---------- the answer ----------

/** An element or the size of an array answer (a one-dimensional array counts as a column). */
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

/** The answer through the adapter (the runtime's fromFj rules: a date to the serial number of its local midnight). */
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
