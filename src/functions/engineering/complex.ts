/**
 * COMPLEX and the IM* functions. Formulas of IMCOS … IMTAN, IMPOWER, IMSQRT ported from formula.js 4.6.1
 * (src/engineering.js), Copyright (c) 2014 Sutoiku, Inc., MIT License (THIRD_PARTY.md); fixed against Excel en-US.
 */
import { gridAt, type FnDef, type Grid } from '../../args.js'
import { textToNumber } from '../../dates.js'
import { atp } from '../date/atp.js'
import { readNumber } from '../../model.js'
import { E, fitNumber, isErr, type CErr, type V } from '../../values.js'

type Unit = '' | 'i' | 'j'

interface Complex {
  re: number
  im: number
  /** The suffix the text had; '' for a real number, which goes with either. */
  unit: Unit
}

/** Excel's largest number as text: larger numbers in complex text are #NUM!. */
const MAX_TEXT = 9.99999999999999e307
/** Excel's limit for the sine and cosine of a real argument. */
const MAX_TRIG = 2 ** 27

const NUMBER = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/

/** Complex text as Excel reads it: spaces ignored, `a`, `bi`, `a±bi` (i or j; b may carry its own sign; `i` is 1i); null if not one. */
function parseComplex(text: string): Complex | null {
  const s = text.replaceAll(' ', '')
  const num = (t: string): [number, string] | null => {
    const m = NUMBER.exec(t)
    if (!m) return null
    const n = readNumber(m[0])
    return Math.abs(n) > MAX_TEXT ? null : [n, t.slice(m[0].length)]
  }
  const isUnit = (u: string): u is 'i' | 'j' => u === 'i' || u === 'j'
  const first = num(s)
  if (!first) {
    const m = /^([+-]?)([ij])$/.exec(s)
    return m ? { re: 0, im: m[1] === '-' ? -1 : 1, unit: m[2] as Unit } : null
  }
  const [a, rest] = first
  if (rest === '') return { re: a, im: 0, unit: '' }
  if (isUnit(rest)) return { re: 0, im: a, unit: rest }
  const sign = rest[0] === '-' ? -1 : rest[0] === '+' ? 1 : 0
  if (!sign) return null
  const tail = rest.slice(1)
  if (isUnit(tail)) return { re: a, im: sign, unit: tail }
  const second = num(tail)
  return second && isUnit(second[1]) ? { re: a, im: sign * second[0], unit: second[1] } : null
}

/**
 * An argument: a number is real, a blank 0, a boolean #VALUE!; text is complex, else a real number as VALUE reads
 * it when no sign follows the first character (`"$1,000"`, `"50%"`, `"18:30"`, `"1/2/2025"`; not `"2025-01-31"`),
 * else #NUM!, as in Excel.
 */
function toComplex(v: V): Complex | CErr {
  if (isErr(v)) return v
  if (typeof v === 'boolean') return E.VALUE
  if (v === null) return { re: 0, im: 0, unit: '' }
  if (typeof v === 'number') return { re: v, im: 0, unit: '' }
  const n = /[+-]/.test(v.trim().slice(1)) ? null : textToNumber(v)
  const z = parseComplex(v) ?? (n === null ? null : { re: n, im: 0, unit: '' as const })
  // −0 is 0 (IMARGUMENT("-1-0i") is π)
  return z ? { re: z.re + 0, im: z.im + 0, unit: z.unit } : E.NUM
}

/** The unit of a result: the arguments' own (i by default); null if they mix i and j. */
function unitOf(zs: readonly Complex[]): 'i' | 'j' | null {
  let u: Unit = ''
  for (const z of zs) {
    if (!z.unit) continue
    if (u && z.unit !== u) return null
    u = z.unit
  }
  return u || 'i'
}

/** A part as Excel writes it in complex text: 15 significant digits, positional up to 21 characters (the General rule of `&` with 20). */
function partText(n: number): string {
  if (n === 0) return '0'
  const [m, e] = Math.abs(n).toExponential(14).split('e')
  const exp = Number(e)
  const digits = m.replace('.', '').replace(/0+$/, '')
  const fixed =
    exp < 0 ? `0.${'0'.repeat(-exp - 1)}${digits}` : digits.length <= exp + 1 ? digits.padEnd(exp + 1, '0') : `${digits.slice(0, exp + 1)}.${digits.slice(exp + 1)}`
  const text = fixed.length <= 21 ? fixed : `${digits[0]}${digits.length > 1 ? '.' + digits.slice(1) : ''}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`
  return n < 0 ? '-' + text : text
}

/** The text of re + im·unit; a part that overflows is #NUM!. */
function complexText(re: number, im: number, unit: 'i' | 'j'): string | CErr {
  const r = fitNumber(re)
  const i = fitNumber(im)
  if (isErr(r)) return r
  if (isErr(i)) return i
  if (i === 0) return partText(r)
  const it = partText(i)
  const imText = it === '1' ? unit : it === '-1' ? `-${unit}` : `${it}${unit}`
  return r === 0 ? imText : `${partText(r)}${i > 0 ? '+' : ''}${imText}`
}

/** a / b by Smith's method; null for b = 0. */
function divide(a: number, b: number, c: number, d: number): [number, number] | null {
  if (c === 0 && d === 0) return null
  if (Math.abs(c) >= Math.abs(d)) {
    const r = d / c
    const den = c + d * r
    return [(a + b * r) / den, (b - a * r) / den]
  }
  const r = c / d
  const den = c * r + d
  return [(a * r + b) / den, (b * r - a) / den]
}

/** sinh and cosh through exp, overflowing at 710 as Excel's do (IMSINH("710") is #NUM!), sinh exact near 0. */
const sh = (x: number) => (Math.expm1(x) - Math.expm1(-x)) / 2
const ch = (x: number) => (Math.exp(x) + Math.exp(-x)) / 2

const sin = (x: number, y: number): [number, number] => [Math.sin(x) * ch(y), Math.cos(x) * sh(y)]
const cos = (x: number, y: number): [number, number] => [Math.cos(x) * ch(y), -Math.sin(x) * sh(y)]
const sinh = (x: number, y: number): [number, number] => [Math.cos(y) * sh(x), Math.sin(y) * ch(x)]
const cosh = (x: number, y: number): [number, number] => [Math.cos(y) * ch(x), Math.sin(y) * sh(x)]

/** a / b with the denominator c² + d² (IMSECH: Excel's last digits). */
function naive(a: number, b: number, c: number, d: number): [number, number] | null {
  const den = c * c + d * d
  return den === 0 ? null : [(a * c + b * d) / den, (b * c - a * d) / den]
}

/** tan z by the double-angle formula, ±i where it overflows; at a pole of that formula, sin z / cos z (Excel: IMTAN("1.5707963267949")). */
function tan(x: number, y: number): [number, number] | null {
  const d = Math.cos(2 * x) + ch(2 * y)
  const s = sh(2 * y)
  if (!Number.isFinite(s) && !Number.isFinite(d)) return [0, s > 0 ? 1 : -1]
  return d === 0 ? divide(...sin(x, y), ...cos(x, y)) : [Math.sin(2 * x) / d, s / d]
}

/** cot z likewise: the double-angle formula, else cos z / sin z (Excel: IMCOT("1e-300"), IMCOT("1+400i")). */
function cot(x: number, y: number): [number, number] | null {
  const d = ch(2 * y) - Math.cos(2 * x)
  const w: [number, number] = [Math.sin(2 * x) / d, -sh(2 * y) / d]
  return w.every(Number.isFinite) ? w : divide(...cos(x, y), ...sin(x, y))
}

/** The argument as formula.js computes it (atan of the quotient, moved by π): IMLOG10's last digits. */
function argument(x: number, y: number): number {
  if (x === 0) return y > 0 ? Math.PI / 2 : -Math.PI / 2
  if (y === 0) return x > 0 ? 0 : Math.PI
  return x > 0 ? Math.atan(y / x) : y >= 0 ? Math.atan(y / x) + Math.PI : Math.atan(y / x) - Math.PI
}

/** r^n by squaring for a whole n, as Excel does (IMPOWER("1+i",100) to the last digit). */
function powerOf(r: number, n: number): number {
  if (!Number.isInteger(n)) return Math.pow(r, n)
  let p = 1
  let b = r
  for (let e = Math.abs(n); e > 0; e = Math.floor(e / 2)) {
    if (e % 2) p *= b
    b *= b
  }
  return n < 0 ? 1 / p : p
}

/** A one-argument function as complex text, #NUM! where it is undefined; `trig` names the part Excel limits to ±2^27. */
function unary(f: (x: number, y: number) => [number, number] | null, trig?: 're' | 'im'): FnDef {
  return {
    args: 'v',
    fn: (v: V) => {
      const z = toComplex(v)
      if (isErr(z)) return z
      if (trig && Math.abs(z[trig]) >= MAX_TRIG) return E.NUM
      const w = f(z.re, z.im)
      return w ? complexText(w[0], w[1], z.unit || 'i') : E.NUM
    },
  }
}

const LN10 = Math.log(10)
const LN2 = Math.log(2)
const log = (x: number, y: number): [number, number] | null => (x === 0 && y === 0 ? null : [Math.log(Math.hypot(x, y)), Math.atan2(y, x)])

const ZERO: Complex = { re: 0, im: 0, unit: '' }

/** The complex numbers of arguments and ranges in order, blank cells as 0; an error or invalid text stops. */
function collect(grids: readonly Grid[]): Complex[] | CErr {
  const out: Complex[] = []
  for (const g of grids) {
    const n = g.rows * g.cols
    const filled = g.h === 0 ? 0 : (g.h - 1) * g.cols + g.w
    for (let i = 0; i < Math.min(n, filled); i++) {
      const z = toComplex(gridAt(g, i))
      if (isErr(z)) return z
      out.push(z)
    }
    // blank cells past the occupied part count once: as a sum or a product they act alike
    if (n > filled) out.push(ZERO)
  }
  return out
}

/** A two-argument operation; mixed units are #NUM!. */
function binary(f: (a: Complex, b: Complex) => [number, number] | null): FnDef {
  return {
    args: 'v v',
    fn: (x: V, y: V) => {
      const a = toComplex(x)
      if (isErr(a)) return a
      const b = toComplex(y)
      if (isErr(b)) return b
      const unit = unitOf([a, b])
      const w = unit && f(a, b)
      return w ? complexText(w[0], w[1], unit!) : E.NUM
    },
  }
}

/** Excel reads the suffix first: a suffix other than i, j or blank is #VALUE! whatever the parts are. */
const SUFFIX = { first: [2], check: { 2: (s: V) => (s === null || s === '' || s === 'i' || s === 'j' ? undefined : E.VALUE) } }

const complex = (re: number, im: number, suffix?: V): V => complexText(re, im, suffix === 'j' ? 'j' : 'i')

function power(v: V, n: number): V {
  const z = toComplex(v)
  if (isErr(z)) return z
  if (z.re === 0 && z.im === 0) return n > 0 ? '0' : E.NUM
  const p = powerOf(Math.sqrt(z.re * z.re + z.im * z.im), n)
  const t = argument(z.re, z.im) * n
  return complexText(p * Math.cos(t), p * Math.sin(t), z.unit || 'i')
}

/** IMSQRT: the modulus as formula.js computes it, underflowing for tiny parts (Excel: IMSQRT("-1e-300") is "0"). */
function sqrt(x: number, y: number): [number, number] {
  if (x === 0 && y === 0) return [0, 0]
  const s = Math.sqrt(Math.sqrt(x * x + y * y))
  const t = argument(x, y) / 2
  return [s * Math.cos(t), s * Math.sin(t)]
}

const part = (f: (z: Complex) => V): FnDef => ({
  args: 'v',
  fn: (v: V) => {
    const z = toComplex(v)
    return isErr(z) ? z : f(z)
  },
})

export const functions: Record<string, FnDef> = {
  COMPLEX: atp('x x v?', complex, SUFFIX),
  IMABS: part((z) => Math.hypot(z.re, z.im)),
  IMAGINARY: part((z) => z.im),
  IMARGUMENT: part((z) => (z.re === 0 && z.im === 0 ? E.DIV0 : Math.atan2(z.im, z.re))),
  IMCONJUGATE: unary((x, y) => [x, -y]),
  IMCOS: unary(cos, 're'),
  IMCOSH: unary(cosh, 'im'),
  IMCOT: unary(cot, 're'),
  IMCSC: unary((x, y) => naive(1, 0, ...sin(x, y)), 're'),
  IMCSCH: unary((x, y) => {
    const d = ch(2 * x) - Math.cos(2 * y)
    return [(2 * sh(x) * Math.cos(y)) / d, (-2 * ch(x) * Math.sin(y)) / d]
  }, 'im'),
  IMDIV: binary((a, b) => divide(a.re, a.im, b.re, b.im)),
  IMEXP: unary((x, y) => [Math.exp(x) * Math.cos(y), Math.exp(x) * Math.sin(y)], 'im'),
  IMLN: unary(log),
  IMLOG10: unary((x, y) => (x === 0 && y === 0 ? null : [Math.log(Math.hypot(x, y)) / LN10, argument(x, y) / LN10])),
  IMLOG2: unary((x, y) => {
    const w = log(x, y)
    return w && [w[0] / LN2, w[1] / LN2]
  }),
  IMPOWER: { args: 'v x', fn: power },
  IMPRODUCT: {
    args: 'G',
    rep: 1,
    fn: (grids: Grid[]) => {
      const zs = collect(grids)
      if (isErr(zs)) return zs
      const unit = unitOf(zs)
      if (!unit) return E.VALUE
      let re = 1
      let im = 0
      for (const z of zs) [re, im] = [re * z.re - im * z.im, re * z.im + im * z.re]
      // Excel: an overflowing product is #VALUE!
      return Number.isFinite(re) && Number.isFinite(im) ? complexText(re, im, unit) : E.VALUE
    },
  },
  IMREAL: part((z) => z.re),
  IMSEC: unary((x, y) => {
    const d = Math.cos(2 * x) + ch(2 * y)
    return [(2 * Math.cos(x) * ch(y)) / d, (2 * Math.sin(x) * sh(y)) / d]
  }, 're'),
  IMSECH: unary((x, y) => naive(1, 0, ...cosh(x, y)), 'im'),
  IMSIN: unary(sin, 're'),
  IMSINH: unary(sinh, 'im'),
  IMSQRT: unary(sqrt),
  IMSUB: binary((a, b) => [a.re - b.re, a.im - b.im]),
  IMSUM: {
    args: 'G',
    rep: 1,
    fn: (grids: Grid[]) => {
      const zs = collect(grids)
      if (isErr(zs)) return zs
      const unit = unitOf(zs)
      if (!unit) return E.VALUE
      return complexText(
        zs.reduce((s, z) => s + z.re, 0),
        zs.reduce((s, z) => s + z.im, 0),
        unit,
      )
    },
  },
  IMTAN: unary(tan, 're'),
}
