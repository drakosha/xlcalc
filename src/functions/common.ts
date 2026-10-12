/**
 * Numeric helpers shared by the function categories.
 *
 * `fact` and `roundDigits` are ported from formula.js 4.6.1 (https://github.com/formulajs/formulajs),
 * src/math-trig.js (FACT, roundBase): Copyright (c) 2014 Sutoiku, Inc., MIT License (full text in
 * THIRD_PARTY.md). Changed: `roundDigits` shifts the decimal point on the shortest decimal
 * representation of any magnitude (formula.js lost the shift for numbers written with an exponent) and
 * leaves a number that is already whole at that many digits unchanged (Excel: ROUND(2.15, 45688) is 2.15).
 */
import { nearZero, zeroed } from '../values.js'

/** Excel's largest argument count of a function newer than Excel 2007 (`_xlfn.NAME` in a file): one fewer than its 255, measured by the `arity` set. */
export const NEWER_MAX_ARGS = 254

/** Excel: sine, cosine and their kin are #NUM! from 2^27 on, for a real argument (IMSIN …) and BESSELJ and BESSELY too. */
export const MAX_TRIG = 2 ** 27

/** Numeric aggregates in input order, with Excel's empty-list results. */
export const sum = (xs: readonly number[]) => xs.reduce((s, x) => s + x, 0)
/** The numbers added in turn: the sum before the last one, and the sum of all. */
function lastAddition(xs: readonly number[]): [before: number, sum: number] {
  let s = 0
  for (let i = 0; i < xs.length - 1; i++) s += xs[i]
  return [s, xs.length ? s + xs[xs.length - 1] : 0]
}
/** SUM: the numbers added in turn, the last addition taking Excel's 0 near the sum before it (`nearZero`). */
export const excelSum = (xs: readonly number[]): number => nearZero(...lastAddition(xs))
/** Whether SUM's last addition left a remainder Excel makes 0 (`zeroed`): AVERAGE is then 0 too. */
export const sumZeroed = (xs: readonly number[]): boolean => zeroed(...lastAddition(xs))
export const sumSquares = (xs: readonly number[]) => xs.reduce((s, x) => s + x * x, 0)
export const product = (xs: readonly number[]) => xs.length ? xs.reduce((p, x) => p * x, 1) : 0
// Do not spread the input into Math.max/min: a whole column can exceed the call argument limit.
export const max = (xs: readonly number[]) => xs.reduce((m, x) => x > m ? x : m, xs.length ? -Infinity : 0)
export const min = (xs: readonly number[]) => xs.reduce((m, x) => x < m ? x : m, xs.length ? Infinity : 0)

const FACTS: number[] = [1, 1]

/** n! for a whole n ≥ 0, by products as formula.js computes it (Infinity past 170). */
export function fact(n: number): number {
  if (n < 0) return NaN
  for (let k = FACTS.length; k <= n && k <= 171; k++) FACTS[k] = FACTS[k - 1] * k
  return n > 171 ? Infinity : FACTS[n]
}

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

/** Digits in radix `r` (letters of either case) as a number; null at a character that is no digit of it. */
export function digitsValue(s: string, r: number): number | null {
  let n = 0
  for (const ch of s) {
    const d = parseInt(ch, 36)
    if (!(d < r)) return null
    n = n * r + d
  }
  return n
}

/** n / s to 15 significant digits, as Excel sees the quotient before CEILING and FLOOR round it (1.5 / 0.1 is 15). */
export const quotient = (n: number, s: number) => Number((n / s).toPrecision(15))
