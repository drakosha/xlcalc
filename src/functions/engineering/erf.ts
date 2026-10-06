/** ERF, ERFC, ERF.PRECISE, ERFC.PRECISE on jStat's erf and erfc (`../jstat.ts`), erfc precise in the tail. */
import type { FnDef } from '../../args.js'
import { erf, erfc } from '../jstat.js'

export const functions: Record<string, FnDef> = {
  // with an upper limit, the difference of the two values (Excel: ERF(5,6) is 1 − ERF(5) to the last bit)
  ERF: { args: 'x x?', fn: (a: number, b?: number) => (b === undefined ? erf(a) : erf(b) - erf(a)) },
  'ERF.PRECISE': { args: 'x', fn: erf },
  ERFC: { args: 'x', fn: erfc },
  'ERFC.PRECISE': { args: 'x', fn: erfc },
}
