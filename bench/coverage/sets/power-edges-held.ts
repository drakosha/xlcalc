/**
 * Held out after `power-neg-held`, whose answers bear out the real root of a negative number where the reciprocal of
 * the power is an odd whole number, `(-8)^0.333333333333333` too, but not `(-8)^0.33333333333333` nor an odd number
 * of 15 digits (`(-2)^(1/999999999999999)`) or of 2^52+1. The rule predicted for the edges: n is 1/y rounded, at most
 * 2^31-1 in magnitude, odd, and 1/n within 2^-48 of y relatively. The cases: 1/(n±d) for n 3, 9 and 101 at distances d
 * on either side of that tolerance, and 1/n for odd n from 10^8 to 10^11 around 2^31.
 * Answers: `test/excel/cases/power-edges-held.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { Probe } from '../probes.js'
import { collector } from './arrays/common.js'

/** The rule predicted: -(-x)^y where 1/y rounds to an odd n of at most 2^31-1 and 1/n is within 2^-48 of y. */
function predict(x: number, y: number): Val {
  const n = Math.round(1 / y)
  const root = Math.abs(n) <= 2 ** 31 - 1 && n % 2 !== 0 && Math.abs(1 / n - y) <= Math.abs(y) * 2 ** -48
  return root ? { t: 'n', v: -Math.exp(y * Math.log(-x)) } : { t: 'e', v: 'NUM' }
}

/** Group, formula, base, power. */
const CASES: [string, string, number, number][] = [
  ...['2E-15', '4E-15', '5E-15', '6E-15', '8E-15', '1E-14', '1.2E-14', '2E-14'].map((d): [string, string, number, number] => ['tolerance', `(-8)^(1/(3+${d}))`, -8, 1 / (3 + Number(d))]),
  ...['4E-15', '6E-15', '8E-15', '1.2E-14'].map((d): [string, string, number, number] => ['tolerance', `(-8)^(1/(3-${d}))`, -8, 1 / (3 - Number(d))]),
  ...['4E-15', '1E-14', '2E-14', '3E-14', '4E-14'].map((d): [string, string, number, number] => ['tolerance', `(-2)^(1/(9+${d}))`, -2, 1 / (9 + Number(d))]),
  ...['2E-13', '3E-13', '4E-13', '6E-13'].map((d): [string, string, number, number] => ['tolerance', `(-2)^(1/(101+${d}))`, -2, 1 / (101 + Number(d))]),
  ...[99999999, 999999999, 2147483647, 2147483649, 4294967297, 99999999999].map((k): [string, string, number, number] => ['size', `(-2)^(1/${k})`, -2, 1 / k]),
  ['size', '(-2)^(-1/2147483647)', -2, -1 / 2147483647],
]

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [group, f] of CASES) add(group, '^', f)
  return out
}

/** The predicted answers by case id. */
export const PREDICTED: Readonly<Record<string, Val>> = Object.fromEntries(buildProbes().map((p, i) => [p.id, predict(CASES[i][2], CASES[i][3])]))

/**
 * The predictions Excel did not confirm (the Mac and Windows alike): 1/y is taken to 15 significant digits, not within
 * 2^-48 (3 + 6E-15 and 9 + 1E-14 are no longer whole, 101 + 4E-13 still is), and 2^31-1 is already too large.
 */
export const MISSED: readonly string[] = ['tolerance#4', 'tolerance#5', 'tolerance#6', 'tolerance#10', 'tolerance#11', 'tolerance#14', 'tolerance#15', 'tolerance#16', 'tolerance#20', 'size#3', 'size#7']
