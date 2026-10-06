/** The argument rule of the former Analysis ToolPak functions (bonds, EDATE, WORKDAY and the like). */
import { toDate, type FnDef } from '../../args.js'
import { E, isErr, toBool, toNum, toText, type CErr, type V } from '../../values.js'

const noBool = (v: V): number | CErr => (typeof v === 'boolean' ? E.VALUE : toNum(v))
const whole = (n: number | CErr) => (isErr(n) ? n : Math.trunc(n))

const SCALAR: Record<string, (v: V) => unknown> = {
  e: (v) => (typeof v === 'boolean' ? E.VALUE : toDate(v)),
  x: noBool,
  j: (v) => whole(noBool(v)),
  n: toNum,
  i: (v) => whole(toNum(v)),
  d: toDate,
  b: toBool,
  s: toText,
}

/**
 * Kinds `e` date, `x` number, `j` integer reject booleans (#VALUE!, as in Excel); scalar kinds are coerced
 * here in argument order, the first error wins; other kinds go through the shared layer.
 */
export function atp(kinds: string, fn: (...a: any[]) => V): FnDef {
  const ks = kinds.split(' ').filter(Boolean)
  const local = ks.map((k) => k[0] in SCALAR)
  return {
    args: ks.map((k, i) => (local[i] ? `a${k.slice(1)}` : k)).join(' '),
    fn: (...a: unknown[]) => {
      for (let i = 0; i < a.length; i++) {
        if (!local[i] || a[i] === undefined) continue
        const v = a[i] as V
        if (isErr(v)) return v
        const c = SCALAR[ks[i][0]](v)
        if (isErr(c)) return c
        a[i] = c
      }
      return fn(...a)
    },
  }
}
