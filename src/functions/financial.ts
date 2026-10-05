/**
 * Financial functions.
 *
 * Ported from formula.js 4.6.1 (https://github.com/formulajs/formulajs), src/financial.js: Copyright (c)
 * 2014 Sutoiku, Inc., MIT License (full text in THIRD_PARTY.md). Changed: arguments are coerced by the
 * engine's argument layer (`src/args.ts`); a zero divisor is #DIV/0! (ISPMT with 0 periods, NPV at a rate
 * of −1); PDURATION needs positive values (#NUM!).
 */
import type { FnDef } from '../args.js'
import { E } from '../values.js'

export const functions: Record<string, FnDef> = {
  ISPMT: { args: 'n i i n', fn: (rate: number, per: number, nper: number, pv: number) => (nper === 0 ? E.DIV0 : pv * rate * (per / nper - 1)) },
  NPV: {
    args: 'n N',
    rep: 1,
    fn: (rate: number, values: number[]) => {
      if (rate === -1) return E.DIV0
      let v = 0
      for (let j = 0; j < values.length; j++) v += values[j] / Math.pow(1 + rate, j + 1)
      return v
    },
  },
  PDURATION: {
    args: 'n n n',
    fn: (rate: number, pv: number, fv: number) => (rate <= 0 || pv <= 0 || fv <= 0 ? E.NUM : (Math.log(fv) - Math.log(pv)) / Math.log(1 + rate)),
  },
}
