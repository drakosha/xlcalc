/**
 * An answer against the reference: numbers within |a−b| ≤ 1e-9·max(1,|a|) (soft: 1e-6), complex numbers as text
 * (a reference with an imaginary unit) by parts within the same tolerance, other strings, booleans and errors
 * (by code) exactly. A mismatch gets a kind.
 */
import type { Val } from './formulajs.js'

export const TOL = 1e-9
export const SOFT = 1e-6

export type Kind =
  | 'number'
  | 'error instead of a value'
  | 'value instead of an error'
  | 'other error code'
  | 'other type'
  | 'string'
  | 'string: case'
  | 'string: number format'
  | 'boolean'
  | 'JS Date'
  | 'NaN or infinity'
  | 'exception'
  | 'hangs'
  | 'other'

export interface Verdict {
  ok: boolean
  /** Matched only within the soft tolerance 1e-6. */
  soft?: boolean
  kind?: Kind
}

const close = (a: number, b: number, tol: number) => Math.abs(a - b) <= tol * Math.max(1, Math.abs(a))

/** Numbers inside text, for `samePin`. */
const NUMBER = /[-+]?(?:\d+\.?\d*|\.\d+)(?:E[-+]?\d+)?/gi

/**
 * The engine's answer against the one a known difference pins: a number within the tolerance Excel's answers have, text
 * with its numbers so, the rest exactly; the last digits Math gives on another Node or platform do not fail a pin.
 */
export function samePin(pinned: Val, got: Val): boolean {
  if (pinned.t !== got.t) return false
  if (pinned.t === 'n') return close(pinned.v, got.v as number, TOL)
  if (pinned.t !== 's' || pinned.v === got.v) return pinned.v === got.v
  const [a, b] = [pinned.v, got.v as string]
  const [na, nb] = [a.match(NUMBER) ?? [], b.match(NUMBER) ?? []]
  return na.length === nb.length && a.split(NUMBER).join('\0') === b.split(NUMBER).join('\0') && na.every((x, i) => close(Number(x), Number(nb[i]), TOL))
}

/** `got` against the reference `ref`; `threw`: the #VALUE! answer came from a formulajs exception. */
export function compare(ref: Val, got: Val, threw?: boolean): Verdict {
  if (got.t === 'other' && (got.v === 'hang' || got.v === 'crash')) return { ok: false, kind: 'hangs' }
  if (got.t === 'throw') return { ok: false, kind: 'exception' }
  if (got.t === 'nan') return { ok: false, kind: 'NaN or infinity' }
  if (got.t === 'date') {
    // bare formulajs: a date as a JS object; its value is compared apart in the report
    return { ok: false, kind: 'JS Date' }
  }
  if (got.t === 'other') return { ok: false, kind: 'other' }
  if (ref.t === 'e') {
    if (got.t !== 'e') return { ok: false, kind: 'value instead of an error' }
    return got.v === ref.v ? { ok: true } : { ok: false, kind: threw ? 'exception' : 'other error code' }
  }
  if (got.t === 'e') return { ok: false, kind: threw ? 'exception' : 'error instead of a value' }
  if (ref.t !== got.t) return { ok: false, kind: 'other type' }
  if (ref.t === 'n') {
    const a = ref.v
    const b = got.v as number
    if (close(a, b, TOL)) return { ok: true }
    if (close(a, b, SOFT)) return { ok: false, soft: true, kind: 'number' }
    return { ok: false, kind: 'number' }
  }
  if (ref.t === 's') {
    const g = got.v as string
    if (g === ref.v || sameComplex(ref.v, g, TOL)) return { ok: true }
    if (g !== ref.v && g.toUpperCase() === ref.v) return { ok: false, kind: 'string: case' }
    return { ok: false, kind: sameNumbers(ref.v, g) ? 'string: number format' : 'string' }
  }
  if (ref.t === 'b') return got.v === ref.v ? { ok: true } : { ok: false, kind: 'boolean' }
  return { ok: false, kind: 'other' }
}

const NUM = '\\d+(?:\\.\\d+)?(?:[eE][+-]?\\d+)?'
/** A complex number as Excel writes it, with an imaginary unit: `3+4i`, `-0.5-2.5j`, `4i`, `-i`, `1E+21+i`. */
const COMPLEX = new RegExp(`^(?:(-?${NUM})([+-])|(-?))(${NUM})?([ij])$`)

function complexParts(s: string): [number, number, string] | null {
  const m = COMPLEX.exec(s)
  if (!m) return null
  const sign = (m[2] ?? m[3]) === '-' ? -1 : 1
  return [m[1] === undefined ? 0 : Number(m[1]), sign * (m[4] === undefined ? 1 : Number(m[4])), m[5]]
}

/**
 * A complex answer (a reference with an imaginary unit) matches by parts within the tolerance and by the unit:
 * the last of 15 significant digits depends on the sine library (Excel for Mac takes Apple's libm, V8 its own).
 * A string without an imaginary unit is compared exactly: numbers as text (TEXT, &, BASE with leading zeros)
 * must match character for character.
 */
function sameComplex(ref: string, got: string, tol: number): boolean {
  const a = complexParts(ref)
  const b = complexParts(got)
  return !!a && !!b && a[2] === b[2] && close(a[0], b[0], tol) && close(a[1], b[1], tol)
}

/** Strings with the same text between their numbers and equal numbers (within the tolerance): "1.5+2i" and "1.50000001+2i". */
function sameNumbers(a: string, b: string): boolean {
  const re = /-?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?/g
  const na = a.match(re) ?? []
  const nb = b.match(re) ?? []
  if (!na.length || na.length !== nb.length || a.replace(re, '#') !== b.replace(re, '#')) return false
  return na.every((x, i) => close(Number(x), Number(nb[i]), TOL))
}

/** A value for the report's tables. */
export function show(v: Val | undefined): string {
  if (!v) return 'none'
  switch (v.t) {
    case 'n':
      return String(Number(v.v.toPrecision(12)))
    case 's':
      return JSON.stringify(v.v)
    case 'b':
      return v.v ? 'TRUE' : 'FALSE'
    case 'e':
      return `#${v.v}`
    case 'date':
      return `Date(${Number(v.v.toPrecision(12))})`
    default:
      return `${v.t}:${v.v}`
  }
}
