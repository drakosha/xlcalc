/**
 * Number sections: digits (`0`, `#`, `?`, the decimal point, the group separator for grouping and for scaling by
 * 1000, `%`), fractions (`# ?/?`, `?/??`, `# ?/8`) and scientific notation (`0.00E+00`, `##0.0E+0`). A section
 * gives text, or null for a number it cannot show (TEXT answers #VALUE!, a cell shows `#`).
 */
import { positional, roundHalfUp } from './general.js'
import type { Ph, Tok } from './parse.js'
import { BLANK, blank, plain, type Style } from './style.js'

/** Digits, decimals and literals of a section by its placeholders. */
function formatDigits(v: number, toks: Tok[], signed: boolean, style: Style): string {
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
  const pct = toks.filter((t) => t.t === 'pct').length
  // a percent or a scale is applied to 15 significant digits, as Excel shows the product (0.285 is 29%)
  let x = pct || scale ? Number(((Math.abs(v) * Math.pow(100, pct)) / Math.pow(1000, scale)).toPrecision(15)) : Math.abs(v)
  const fracPh = dot < 0 ? [] : toks.slice(dot + 1).filter((t): t is Ph => t.t === 'ph')
  const intPh = toks.slice(0, intEnd).filter((t): t is Ph => t.t === 'ph')
  const fixed = positional(x, fracPh.length)
  x = Number(fixed)
  let [ip, fp = ''] = fixed.split('.')
  const zeros = intPh.filter((p) => p.c === '0').length
  if (ip === '0' && zeros === 0) ip = ''
  ip = ip.padStart(zeros, '0')
  if (grouping) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, style.group)
  // integer part: right to left over digit placeholders, leftover digits go before the first one
  let out = ''
  let rest = ip
  const firstPh = toks.findIndex((t, i) => i < intEnd && t.t === 'ph')
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
    for (const t of toks.slice(dot + 1)) {
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
 * A fraction: placeholders right before `/` are the numerator, placeholders before them (if any) the integer
 * part; after `/` either placeholders (the denominator has at most that many digits, the closest fraction is
 * taken) or a fixed denominator (`?/8`). With an integer part a negative number keeps its minus even when it
 * shows 0, and a whole 0 shows `0` in a `#` place (`-0    `).
 */
function formatFraction(v: number, toks: Tok[], slash: number, signed: boolean, style: Style): string | null {
  let i = slash - 1
  while (i >= 0 && toks[i].t === 'ph') i--
  const numPh = toks.slice(i + 1, slash) as Ph[]
  const intToks = toks.slice(0, i + 1)
  const hasInt = intToks.some((t) => t.t === 'ph')
  let j = slash + 1
  let fixed = 0
  const denPh: Ph[] = []
  if (isDigitLit(toks[j])) {
    let digits = ''
    for (; j < toks.length && (isDigitLit(toks[j]) || (toks[j].t === 'ph' && (toks[j] as Ph).c === '0')); j++) digits += toks[j].t === 'lit' ? (toks[j] as { s: string }).s : '0'
    fixed = Number(digits)
  } else for (; j < toks.length && toks[j].t === 'ph'; j++) denPh.push(toks[j] as Ph)
  if (!numPh.length || (!fixed && !denPh.length)) return null
  const x = Math.abs(v)
  let ip = hasInt ? Math.floor(x) : 0
  const frac = x - ip
  let num = 0
  let den = fixed
  if (fixed) num = Math.round(frac * fixed)
  else {
    let best = Infinity
    for (let d = 1, max = 10 ** denPh.length - 1; d <= max; d++) {
      const n = Math.round(frac * d)
      const err = Math.abs(frac - n / d)
      if (err < best - 1e-12) [best, num, den] = [err, n, d]
    }
  }
  if (hasInt && num === den) [ip, num] = [ip + 1, 0]
  if (num > MAX_NUMERATOR) return null
  // unused digit places: `0` gives a zero before the digits, `?` a space (after the denominator's digits)
  const pad = (s: string, ph: Ph[], left: boolean) => {
    const free = ph.slice(0, Math.max(0, ph.length - s.length))
    const zeros = '0'.repeat(free.filter((t) => t.c === '0').length)
    const spaces = blank(style).repeat(free.filter((t) => t.c === '?').length)
    return left ? spaces + zeros + s : zeros + s + spaces
  }
  const last = intToks.findLastIndex((t) => t.t === 'ph')
  const wholeZero = ip === 0 && num === 0
  const intPart = hasInt ? formatDigits(ip, wholeZero ? intToks.map((t, k): Tok => (k === last ? { t: 'ph', c: '0' } : t)) : intToks, false, style) : intToks.map((t) => plain(t, style)).join('')
  const denText = fixed ? String(fixed) : pad(String(den), denPh, false)
  const fracText = pad(String(num), numPh, true) + '/' + denText
  const tail = toks
    .slice(j)
    .map((t) => plain(t, style))
    .join('')
  // a whole number: the fraction's place is blank (laid out for AutoFit, its digits' room and its slash)
  const body = hasInt && num === 0 ? intPart + (style.layout ? fracText.replace(/[^/]/g, BLANK) : ' '.repeat(fracText.length)) : intPart + fracText
  return (signed && v < 0 && (hasInt || num !== 0) ? '-' : '') + body + tail
}

/**
 * Scientific notation: the mantissa has as many integer digits as the integer placeholders allow, the exponent
 * is a multiple of their count (`##0.0E+0` shows 500.0E-3), padded to its placeholders; zero shows every
 * integer placeholder as 0.
 */
function formatExponent(v: number, toks: Tok[], ex: number, signed: boolean, style: Style): string {
  const mant = toks.slice(0, ex)
  const dot = mant.findIndex((t) => t.t === 'dot')
  const width = Math.max(1, mant.slice(0, dot < 0 ? mant.length : dot).filter((t) => t.t === 'ph').length)
  const decimals = dot < 0 ? 0 : mant.slice(dot + 1).filter((t) => t.t === 'ph').length
  const a = Math.abs(v)
  let e = 0
  let m = 0
  if (a !== 0) {
    const [d, x] = a.toExponential(14).split('e')
    const e0 = Number(x)
    e = Math.floor(e0 / width) * width
    m = Number(`${d}e${e0 - e}`)
    if (Number(roundHalfUp(m, decimals)) >= 10 ** width) [e, m] = [e + width, Number(`${d}e${e0 - e - width}`)]
  }
  const body = formatDigits(m, a === 0 ? mant.map((t): Tok => (t.t === 'ph' ? { t: 'ph', c: '0' } : t)) : mant, false, style)
  const plus = (toks[ex] as Extract<Tok, { t: 'exp' }>).plus
  const after = toks.slice(ex + 1)
  const places = after.filter((t) => t.t === 'ph').length
  const digits = String(Math.abs(e)).padStart(places, '0')
  const rest = after.filter((t) => t.t !== 'ph').map((t) => plain(t, style))
  return (signed && v < 0 ? '-' : '') + body + 'E' + (e < 0 ? '-' : plus ? '+' : '') + digits + rest.join('')
}

/** A number by a section of digits, a fraction or scientific notation; null when it cannot be shown. */
export function formatNumber(v: number, toks: Tok[], signed: boolean, style: Style): string | null {
  const ex = toks.findIndex((t) => t.t === 'exp')
  if (ex >= 0) return formatExponent(v, toks, ex, signed, style)
  const slash = toks.findIndex((t, i) => t.t === 'slash' && toks[i - 1]?.t === 'ph')
  if (slash >= 0) return formatFraction(v, toks, slash, signed, style)
  return formatDigits(v, toks, signed, style)
}
