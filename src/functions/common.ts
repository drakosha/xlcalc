/**
 * Numeric helpers shared by the function categories.
 *
 * `roundDigits` is ported from formula.js 4.6.1 (https://github.com/formulajs/formulajs), src/math-trig.js
 * (roundBase): Copyright (c) 2014 Sutoiku, Inc., MIT License (full text in THIRD_PARTY.md). Changed: it
 * shifts the decimal point on the shortest decimal representation of any magnitude (formula.js lost the
 * shift for numbers written with an exponent) and leaves a number that is already whole at that many digits
 * unchanged (Excel: ROUND(2.15, 45688) is 2.15).
 */

/** x with the decimal point moved by k places, computed on its shortest decimal representation (1.005 → 100.5). */
function shift(x: number, k: number): number {
  const [m, e] = x.toExponential().split('e')
  return Number(`${m}e${Number(e) + k}`)
}

/** ROUND, ROUNDUP, ROUNDDOWN and TRUNC: `f` applied to |x| at `digits` decimal places, symmetric around zero. */
export function roundDigits(x: number, digits: number, f: (y: number) => number): number {
  if (x === 0) return 0
  const y = shift(Math.abs(x), digits)
  if (!Number.isFinite(y) || y >= 2 ** 53) return x
  return Math.sign(x) * shift(f(y), -digits)
}
