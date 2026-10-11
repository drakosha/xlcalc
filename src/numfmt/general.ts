/**
 * The General format: a number as text as Excel converts it (`&`, CONCAT), as TEXT's General writes it (at most
 * 11 characters), and as a cell shows it in a column of a given width (`test/excel/display/general.json`).
 */
import { roundDigits } from '../functions/common.js'
import { DIGIT, glyph, hashes, textWidth } from './glyphs.js'

/** Round half away from zero on the decimal representation (1.005 → 1.01). */
export function roundHalfUp(abs: number, k: number): string {
  return roundDigits(abs, k, Math.round).toFixed(k)
}

/** `abs` with `k` decimals, rounded half away from zero; from 10^21 on, its 15 significant digits and zeros. */
export function positional(abs: number, k: number): string {
  if (abs < 1e21) return roundHalfUp(abs, k)
  const [m, e] = abs.toExponential(14).split('e')
  return m.replace('.', '').padEnd(Number(e) + 1, '0') + (k ? '.' + '0'.repeat(k) : '')
}

/** Trailing zeros of the decimals dropped, and the point when none is left. */
const strip = (s: string) => (s.includes('.') ? s.replace(/\.?0+$/, '') : s)

/** The exponent of scientific notation as Excel writes it: its sign and two digits at least (`E+20`, `E-06`). */
const power = (e: number) => `E${e < 0 ? '-' : '+'}${String(Math.abs(e)).padStart(2, '0')}`

/** The decimal exponent of a positive number rounded to 15 significant digits. */
const exponent = (a: number) => Number(a.toExponential(14).split('e')[1])

/** Scientific notation with `m` decimals of the mantissa at most (trailing zeros dropped), an exponent of two digits at least. */
function sci(a: number, m: number): string {
  const [digits, ex] = a.toExponential(14).split('e')
  let e = Number(ex)
  let mant = roundHalfUp(Number(digits), m)
  if (Number(mant) >= 10) [e, mant] = [e + 1, roundHalfUp(Number(mant) / 10, m)]
  return strip(mant) + power(e)
}

/**
 * A number as text the way Excel converts one (`&`, CONCAT, text arguments): rounded to 15 significant
 * digits, in positional notation unless that takes more than `max` characters (the sign aside), else in
 * scientific notation with an exponent of two digits at least (`1E+20`, `7.4500104300146E-06`). The formula bar
 * takes 21 characters (`100000000000000000000`), `&` 20.
 */
export function generalText(n: number, max = 20): string {
  if (n === 0) return '0'
  const [m, e] = Math.abs(n).toExponential(14).split('e')
  const exp = Number(e)
  const digits = m.replace('.', '').replace(/0+$/, '')
  const fixed =
    exp < 0 ? `0.${'0'.repeat(-exp - 1)}${digits}` : digits.length <= exp + 1 ? digits.padEnd(exp + 1, '0') : `${digits.slice(0, exp + 1)}.${digits.slice(exp + 1)}`
  const text = fixed.length <= max ? fixed : `${digits[0]}${digits.length > 1 ? '.' + digits.slice(1) : ''}${power(exp)}`
  return n < 0 ? '-' + text : text
}

/**
 * A number in the General format of TEXT: at most 11 characters, the sign aside. Positional when that
 * shows all of the number's 15 significant digits or as many as scientific notation would (a fraction
 * keeps up to 9 decimals, `0.000123457`), else scientific with the mantissa rounded to fit
 * (`1.23457E-05`, `1E+11`).
 */
export function generalFormat(n: number): string {
  if (n === 0) return '0'
  const a = Math.abs(n)
  // the shortest decimal text of 11 characters at most is the General text (most values of a book)
  const short = String(a)
  if (short.length <= 11 && !short.includes('e')) return n < 0 ? '-' + short : short
  const [m, ex] = a.toExponential(14).split('e')
  const e = Number(ex)
  const need = m.replace('.', '').replace(/0+$/, '').length
  const sciCap = 11 - 3 - Math.max(2, String(Math.abs(e)).length)
  const fixedCap = e >= 0 ? (e < 11 ? Math.max(10, e + 1) : 0) : 10 + e
  let text: string | null = null
  if (fixedCap > 0 && (need <= fixedCap || fixedCap >= sciCap)) {
    const fixed = strip(roundHalfUp(a, e >= 0 ? Math.max(0, 9 - e) : 9))
    if (fixed.length <= 11) text = fixed
  }
  if (text === null) {
    const [sm, se] = a.toExponential(sciCap - 1).split('e')
    text = strip(sm) + power(Number(se))
  }
  return n < 0 ? '-' + text : text
}

/**
 * A number in the General format as a cell of `space` points shows it (decimal point `.`), as measured in
 * Excel: the General text when its characters fit, counting each as a digit and a leading minus by its width;
 * else, for a number from 10^-4 to under 10^11, fewer decimals that still show a digit of the number; else
 * scientific notation with fewer digits; else `0` for a number that rounds to it in 1 to 4 digits' room; a
 * positive number from 10^-4 whose rounding is one digit shows it in any width; the shortest scientific form
 * when its own width fits and the number rounds to 6 digits or more, or is under 10^-4; else `#` across the cell.
 */
export function generalFit(n: number, space: number): string {
  if (n === 0) return '0'
  const neg = n < 0
  const a = Math.abs(n)
  const sign = neg ? '-' : ''
  const cells = Math.floor((space - (neg ? glyph('-') : 0)) / DIGIT)
  const full = generalFormat(a)
  if (full.length <= cells) return sign + full
  const e = exponent(a)
  const positional = e >= -4 && e <= 10
  // more decimals than the room left by the integer part and the point cannot fit
  if (positional)
    for (let d = Math.max(0, Math.min(9, cells - Math.max(1, e + 1) - 1)); d >= 0; d--) {
      if (d > 0 && a < 10 ** -d) continue
      const t = strip(roundHalfUp(a, d))
      if (t !== '0' && t.length <= cells) return sign + t
    }
  for (let m = 5; m >= 0; m--) {
    const t = sci(a, m)
    if (t.length <= cells) return sign + t
  }
  const whole = roundHalfUp(a, 0)
  if (whole === '0' && cells >= 1 && cells < 5) return sign + whole
  if (!neg && positional && cells < 1 && whole.length === 1) return whole
  const shortest = sign + sci(a, 0)
  if ((whole.length > 5 || e < -4) && textWidth(shortest) <= space) return shortest
  return hashes(space, true)
}
