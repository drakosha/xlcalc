/**
 * ERF, ERFC, ERF.PRECISE, ERFC.PRECISE. The erfc fit ported from jStat 1.9.6 (src/special.js), Copyright (c) 2013
 * jStat, MIT License (THIRD_PARTY.md), kept apart so that erfc stays precise in the tail.
 */
const COF = [
  -1.3026537197817094, 6.4196979235649026e-1, 1.9476473204185836e-2, -9.561514786808631e-3, -9.46595344482036e-4, 3.66839497852761e-4, 4.2523324806907e-5, -2.0278578112534e-5,
  -1.624290004647e-6, 1.30365583558e-6, 1.5626441722e-8, -8.5238095915e-8, 6.529054439e-9, 5.059343495e-9, -9.91364156e-10, -2.27365122e-10, 9.6467911e-11, 2.394038e-12, -6.886027e-12,
  8.94487e-13, 3.13092e-13, -1.12708e-13, 3.81e-16, 7.106e-15, -1.523e-15, -9.4e-17, 1.21e-16, -2.8e-17,
]

// TODO: import erfc from ../jstat.ts once the statistics branch exports it, and drop this copy
/** erfc(x) for x ≥ 0. */
function erfcPos(x: number): number {
  const t = 2 / (2 + x)
  const ty = 4 * t - 2
  let d = 0
  let dd = 0
  for (let j = COF.length - 1; j > 0; j--) {
    const tmp = d
    d = ty * d - dd + COF[j]
    dd = tmp
  }
  return t * Math.exp(-x * x + 0.5 * (COF[0] + ty * d) - dd)
}

const erf = (x: number) => (x < 0 ? erfcPos(-x) - 1 : 1 - erfcPos(x))
const erfc = (x: number) => (x < 0 ? 2 - erfcPos(-x) : erfcPos(x))

import type { FnDef } from '../../args.js'

export const functions: Record<string, FnDef> = {
  // with an upper limit, the difference of the two values (Excel: ERF(5,6) is 1 − ERF(5) to the last bit)
  ERF: { args: 'x x?', fn: (a: number, b?: number) => (b === undefined ? erf(a) : erf(b) - erf(a)) },
  'ERF.PRECISE': { args: 'x', fn: erf },
  ERFC: { args: 'x', fn: erfc },
  'ERFC.PRECISE': { args: 'x', fn: erfc },
}
