/**
 * Number sections: digits (`0`, `#`, `?`, the decimal point, the group separator for grouping and for scaling by
 * 1000, `%`), fractions (`# ?/?`, `?/??`, `# ?/8`) and scientific notation (`0.00E+00`, `##0.0E+0`). A section
 * gives text, or null for a number it cannot show (TEXT answers #VALUE!, a cell shows `#`).
 */
import { positional, roundHalfUp, sig15 } from './general.js'
import type { Ph, Tok } from './parse.js'
import { BLANK, blank, plain, type Style } from './style.js'

/** What the digits of a section's tokens give whatever the value (`formatDigits`). */
interface Digits {
  dot: number
  intEnd: number
  /** The powers of 1000 its commas scale by. */
  scale: number
  grouping: boolean
  pct: number
  fracPh: Ph[]
  /** The integer part's `0` placeholders. */
  zeros: number
  firstPh: number
}

/** Plans by a section's tokens, or by the parts of them a fraction and scientific notation keep (`Plan`). */
const DIGITS = new WeakMap<Tok[], Digits>()

function digitsOf(toks: Tok[]): Digits {
  let p = DIGITS.get(toks)
  if (p) return p
  const dot = toks.findIndex((t) => t.t === 'dot')
  const intEnd = dot < 0 ? toks.length : dot
  let scale = 0
  let grouping = false
  for (let i = 0; i < intEnd; i++) {
    if (toks[i].t !== 'comma') continue
    const before = toks.slice(0, i).some((t) => t.t === 'ph')
    const after = toks.slice(i + 1, intEnd).some((t) => t.t === 'ph')
    if (before && after) grouping = true
    else if (before) scale++
  }
  // commas after the decimal digits scale too (`0.0,,` shows millions)
  for (let i = intEnd + 1; i < toks.length; i++) if (toks[i].t === 'comma' && !toks.slice(i + 1).some((t) => t.t === 'ph')) scale++
  p = {
    dot,
    intEnd,
    scale,
    grouping,
    pct: toks.filter((t) => t.t === 'pct').length,
    fracPh: dot < 0 ? [] : toks.slice(dot + 1).filter((t): t is Ph => t.t === 'ph'),
    zeros: toks.slice(0, intEnd).filter((t) => t.t === 'ph' && t.c === '0').length,
    firstPh: toks.findIndex((t, i) => i < intEnd && t.t === 'ph'),
  }
  DIGITS.set(toks, p)
  return p
}

/** Digits, decimals and literals of a section by its placeholders. */
function formatDigits(v: number, toks: Tok[], signed: boolean, style: Style): string {
  const { dot, intEnd, scale, grouping, pct, fracPh, zeros, firstPh } = digitsOf(toks)
  // a percent or a scale is applied to 15 significant digits, as Excel shows the product (0.285 is 29%)
  let x = pct || scale ? Number(((Math.abs(v) * Math.pow(100, pct)) / Math.pow(1000, scale)).toPrecision(15)) : Math.abs(v)
  const fixed = positional(x, fracPh.length)
  x = Number(fixed)
  let [ip, fp = ''] = fixed.split('.')
  if (ip === '0' && zeros === 0) ip = ''
  ip = ip.padStart(zeros, '0')
  if (grouping) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, style.group)
  // integer part: right to left over digit placeholders, leftover digits go before the first one
  let out = ''
  let rest = ip
  for (let i = intEnd - 1; i >= 0; i--) {
    const t = toks[i]
    if (t.t === 'ph') {
      const d = rest.length ? rest[rest.length - 1] : t.c === '0' ? '0' : t.c === '?' ? blank(style) : ''
      rest = rest.slice(0, -1)
      out = (i === firstPh ? rest : '') + d + out
      if (i === firstPh) rest = ''
    } else if (t.t !== 'comma') out = plain(t, style) + out
  }
  if (firstPh < 0) out = ip + out
  // fractional part: trailing `#` drops zeros, `?` gives a space
  if (dot >= 0) {
    const digits = fp.split('')
    for (let k = fracPh.length - 1; k >= 0 && digits[k] === '0' && fracPh[k].c !== '0'; k--) digits[k] = fracPh[k].c === '?' ? blank(style) : ''
    let k = 0
    out += style.decimal
    for (let i = dot + 1; i < toks.length; i++) {
      const t = toks[i]
      if (t.t === 'ph') out += digits[k++] ?? ''
      else if (t.t !== 'comma') out += plain(t, style)
    }
  }
  const negative = signed && v < 0 && x !== 0
  return (negative ? '-' : '') + out
}

const isDigitLit = (t: Tok | undefined) => t?.t === 'lit' && /^[0-9]$/.test(t.s)

/** Excel does not show an improper fraction whose numerator is past 2^31 − 2 (2,147,483,647 is `#`). */
const MAX_NUMERATOR = 2 ** 31 - 2

/**
 * The fraction num/den with den from 1 to `max` that a fraction format shows for `x` ≥ 0: going up the denominators,
 * each one closer to `x` than the last taken by more than 1e-12 is taken. Such a denominator is closer than every smaller
 * one, so it is a convergent or an intermediate fraction of `x`'s continued fraction; only those are tried, in increasing
 * order, and within a run of intermediate fractions (the error falls as they grow) the next taken one is found by halving.
 */
export function closestFraction(x: number, max: number): [number, number] {
  let [best, num, den] = [Infinity, 0, 1]
  const beats = (d: number) => Math.abs(x - Math.round(x * d) / d) < best - 1e-12
  const take = (d: number) => {
    const n = Math.round(x * d)
    ;[best, num, den] = [Math.abs(x - n / d), n, d]
  }
  take(1)
  // the continued fraction of the double itself: x = m / 2^e exactly, its terms by Euclid's algorithm on integers
  let [m, e] = [x, 0]
  while (!Number.isInteger(m)) [m, e] = [m * 2, e + 1]
  const whole = 1n << BigInt(e)
  let [p, q] = [whole, BigInt(m) % whole]
  // the denominators of run k are j·q_k + q_(k-1) for j = 1 to a_(k+1), the last of them q_(k+1); with no error left above 1e-12 none is taken
  for (let [before, qk] = [0, 1]; q !== 0n && best > 1e-12; ) {
    const a = p / q
    ;[p, q] = [q, p % q]
    const terms = Number(a)
    const top = Math.min(terms, Math.floor((max - before) / qk))
    const d = (j: number) => j * qk + before
    for (let j = 1; j <= top && best > 1e-12 && beats(d(top)); j++) {
      if (!beats(d(j))) {
        let hi = top
        for (let lo = j + 1; lo < hi; ) {
          const mid = Math.floor((lo + hi) / 2)
          if (beats(d(mid))) hi = mid
          else lo = mid + 1
        }
        j = hi
      }
      take(d(j))
    }
    if (!(top >= terms)) break
    ;[before, qk] = [qk, d(terms)]
  }
  return [num, den]
}

/** A fraction's parts by a section's tokens (`formatFraction`); null where it has no numerator or no denominator. */
interface Fraction {
  numPh: Ph[]
  /** The tokens before the numerator, and the same with a `0` for their last placeholder (a whole 0). */
  intToks: Tok[]
  intZero: Tok[]
  hasInt: boolean
  /** A fixed denominator (`?/8`), else 0 and the denominator's placeholders. */
  fixed: number
  denPh: Ph[]
  tail: Tok[]
}

function fractionOf(toks: Tok[], slash: number): Fraction | null {
  let i = slash - 1
  while (i >= 0 && toks[i].t === 'ph') i--
  const numPh = toks.slice(i + 1, slash) as Ph[]
  const intToks = toks.slice(0, i + 1)
  let j = slash + 1
  let fixed = 0
  const denPh: Ph[] = []
  if (isDigitLit(toks[j])) {
    let digits = ''
    for (; j < toks.length && (isDigitLit(toks[j]) || (toks[j].t === 'ph' && (toks[j] as Ph).c === '0')); j++) digits += toks[j].t === 'lit' ? (toks[j] as { s: string }).s : '0'
    fixed = Number(digits)
  } else for (; j < toks.length && toks[j].t === 'ph'; j++) denPh.push(toks[j] as Ph)
  if (!numPh.length || (!fixed && !denPh.length)) return null
  const last = intToks.findLastIndex((t) => t.t === 'ph')
  const intZero = intToks.map((t, k): Tok => (k === last ? { t: 'ph', c: '0' } : t))
  return { numPh, intToks, intZero, hasInt: last >= 0, fixed, denPh, tail: toks.slice(j) }
}

/**
 * A fraction: placeholders right before `/` are the numerator, placeholders before them (if any) the integer
 * part; after `/` either placeholders (the denominator has at most that many digits, the closest fraction is
 * taken) or a fixed denominator (`?/8`). With an integer part a negative number keeps its minus even when it
 * shows 0, and a whole 0 shows `0` in a `#` place (`-0    `).
 */
function formatFraction(v: number, f: Fraction, signed: boolean, style: Style): string | null {
  const { numPh, intToks, intZero, hasInt, fixed, denPh } = f
  const x = Math.abs(v)
  let ip = hasInt ? Math.floor(x) : 0
  const frac = x - ip
  let num = 0
  let den = fixed
  if (fixed) num = Math.round(frac * fixed)
  else [num, den] = closestFraction(frac, 10 ** denPh.length - 1)
  if (hasInt && num === den) [ip, num] = [ip + 1, 0]
  if (num > MAX_NUMERATOR) return null
  // unused digit places: `0` gives a zero before the digits, `?` a space (after the denominator's digits)
  const pad = (s: string, ph: Ph[], left: boolean) => {
    const free = ph.slice(0, Math.max(0, ph.length - s.length))
    const zeros = '0'.repeat(free.filter((t) => t.c === '0').length)
    const spaces = blank(style).repeat(free.filter((t) => t.c === '?').length)
    return left ? spaces + zeros + s : zeros + s + spaces
  }
  const wholeZero = ip === 0 && num === 0
  const intPart = hasInt ? formatDigits(ip, wholeZero ? intZero : intToks, false, style) : intToks.map((t) => plain(t, style)).join('')
  const denText = fixed ? String(fixed) : pad(String(den), denPh, false)
  const fracText = pad(String(num), numPh, true) + '/' + denText
  const tail = f.tail.map((t) => plain(t, style)).join('')
  // a whole number: the fraction's place is blank (laid out for AutoFit, its digits' room and its slash)
  const body = hasInt && num === 0 ? intPart + (style.layout ? fracText.replace(/[^/]/g, BLANK) : ' '.repeat(fracText.length)) : intPart + fracText
  return (signed && v < 0 && (hasInt || num !== 0) ? '-' : '') + body + tail
}

/** Scientific notation's parts by a section's tokens (`formatExponent`). */
interface Exponent {
  /** The mantissa's tokens, and the same with each placeholder a `0` (zero shows them so). */
  mant: Tok[]
  mantZero: Tok[]
  /** The mantissa's integer placeholders (one at least) and decimals. */
  width: number
  decimals: number
  plus: boolean
  /** The exponent's placeholders, and the tokens after it that are none. */
  places: number
  rest: Tok[]
}

function exponentOf(toks: Tok[], ex: number): Exponent {
  const mant = toks.slice(0, ex)
  const dot = mant.findIndex((t) => t.t === 'dot')
  const after = toks.slice(ex + 1)
  return {
    mant,
    mantZero: mant.map((t): Tok => (t.t === 'ph' ? { t: 'ph', c: '0' } : t)),
    width: Math.max(1, mant.slice(0, dot < 0 ? mant.length : dot).filter((t) => t.t === 'ph').length),
    decimals: dot < 0 ? 0 : mant.slice(dot + 1).filter((t) => t.t === 'ph').length,
    plus: (toks[ex] as Extract<Tok, { t: 'exp' }>).plus,
    places: after.filter((t) => t.t === 'ph').length,
    rest: after.filter((t) => t.t !== 'ph'),
  }
}

/**
 * Scientific notation: the mantissa has as many integer digits as the integer placeholders allow, the exponent
 * is a multiple of their count (`##0.0E+0` shows 500.0E-3), padded to its placeholders; zero shows every
 * integer placeholder as 0.
 */
function formatExponent(v: number, x: Exponent, signed: boolean, style: Style): string {
  const { width, decimals } = x
  const a = Math.abs(v)
  let e = 0
  let m = 0
  if (a !== 0) {
    const [d, e0] = sig15(a)
    e = Math.floor(e0 / width) * width
    m = Number(`${d}e${e0 - e}`)
    if (Number(roundHalfUp(m, decimals)) >= 10 ** width) [e, m] = [e + width, Number(`${d}e${e0 - e - width}`)]
  }
  const body = formatDigits(m, a === 0 ? x.mantZero : x.mant, false, style)
  const digits = String(Math.abs(e)).padStart(x.places, '0')
  const rest = x.rest.map((t) => plain(t, style))
  return (signed && v < 0 ? '-' : '') + body + 'E' + (e < 0 ? '-' : x.plus ? '+' : '') + digits + rest.join('')
}

/** What a section of a number is, its parts kept: the token arrays a fraction and scientific notation pass on are then the same each time, and their digits' plans are found again. */
type Plan = { exponent: Exponent } | { fraction: Fraction | null } | null

/** Plans by a section's tokens, which `parseFormat` keeps one array of per code. */
const PLANS = new WeakMap<Tok[], Plan>()

function planOf(toks: Tok[]): Plan {
  let p = PLANS.get(toks)
  if (p !== undefined) return p
  const ex = toks.findIndex((t) => t.t === 'exp')
  const slash = toks.findIndex((t, i) => t.t === 'slash' && toks[i - 1]?.t === 'ph')
  p = ex >= 0 ? { exponent: exponentOf(toks, ex) } : slash >= 0 ? { fraction: fractionOf(toks, slash) } : null
  PLANS.set(toks, p)
  return p
}

/** A number by a section of digits, a fraction or scientific notation; null when it cannot be shown. */
export function formatNumber(v: number, toks: Tok[], signed: boolean, style: Style): string | null {
  const p = planOf(toks)
  if (!p) return formatDigits(v, toks, signed, style)
  if ('exponent' in p) return formatExponent(v, p.exponent, signed, style)
  return p.fraction && formatFraction(v, p.fraction, signed, style)
}
