/**
 * Held out after `power-edges-held`, whose answers give the real root of a negative number for 1/y of 999999999 and
 * `#NUM!` for 2147483647 (2^31-1). The rule predicted for the bound: 1/y rounded to a whole number as a 32-bit
 * integer would round it, |1/y| + 0.5 below 2^31-1, so 2147483645 is still a root and 2^30 is no bound.
 * Answers: `test/excel/cases/power-bound-held.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import type { Probe } from '../probes.js'
import { collector } from './arrays/common.js'

/** Odd whole numbers n, and the sign of the power 1/n. */
const CASES: [number, number][] = [[1000000001, 1], [1073741823, 1], [1073741825, 1], [1500000001, 1], [2147483643, 1], [2147483645, 1], [2147483645, -1]]

const powerOf = ([k, sign]: [number, number]) => sign / k

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [k, sign] of CASES) add('size', '^', `(-2)^(${sign < 0 ? '-' : ''}1/${k})`)
  return out
}

/** The predicted answers by case id: -2^y where |1/y| + 0.5 is below 2^31-1. */
export const PREDICTED: Readonly<Record<string, Val>> = Object.fromEntries(
  buildProbes().map((p, i) => {
    const y = powerOf(CASES[i])
    return [p.id, Math.abs(1 / y) + 0.5 < 2 ** 31 - 1 ? { t: 'n', v: -Math.exp(y * Math.log(2)) } : { t: 'e', v: 'NUM' }]
  }),
)
