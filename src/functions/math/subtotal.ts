/**
 * SUBTOTAL and AGGREGATE on the statistical functions. A cell whose formula calls either of them (anywhere in the
 * formula) is skipped as nested, as in Excel; there are no hidden rows, so codes 101–111 equal 1–11 and AGGREGATE's
 * options 1, 3, 5, 7 equal 0, 2, 4, 6. This engine's own.
 */
import { E, isErr, Rng, toNum, type CErr, type V } from '../../values.js'
import type { Cx, FnDef, Grid } from '../../args.js'
import { functions as stat } from '../statistical/descriptive.js'
import { functions as sums } from './sums.js'

/** A statistical function applied to the numbers of the references, with its rule for no numbers. */
function apply(name: string, xs: number[], k?: number): V {
  const def = name === 'SUM' || name === 'PRODUCT' ? sums[name] : stat[name]
  if (def.empty && !xs.length) return E[def.empty]
  if (name === 'MODE.SNGL') return def.fn([{ rows: 1, cols: xs.length, h: 1, w: xs.length, v: xs } satisfies Grid])
  return k === undefined ? def.fn(xs) : def.fn(xs, k)
}

const NAMES = ['AVERAGE', 'COUNT', 'COUNTA', 'MAX', 'MIN', 'PRODUCT', 'STDEV.S', 'STDEV.P', 'SUM', 'VAR.S', 'VAR.P', 'MEDIAN', 'MODE.SNGL', 'LARGE', 'SMALL', 'PERCENTILE.INC', 'QUARTILE.INC', 'PERCENTILE.EXC', 'QUARTILE.EXC']

/**
 * The function `code` (1–13) over the references' values: COUNT counts numbers, COUNTA every non-blank cell
 * (errors too), the rest take the numbers and stop at the first error, unless errors are skipped.
 */
function over(cx: Cx, code: number, refs: readonly (V | Rng)[], skipNested: boolean, skipErrors: boolean, k?: number): V {
  const xs: number[] = []
  let filled = 0
  let first: CErr | undefined
  for (const x of refs) {
    // only references: a value in their place is #VALUE!
    if (!(x instanceof Rng)) return isErr(x) ? x : E.VALUE
    for (const v of skipNested ? cx.flatOuter(x) : cx.flat(x)) {
      if (v === null) continue
      if (isErr(v)) {
        if (skipErrors) continue
        first ??= v
      }
      filled++
      if (typeof v === 'number') xs.push(v)
    }
  }
  if (code === 2) return xs.length
  if (code === 3) return filled
  return first ?? apply(NAMES[code - 1], xs, k)
}

function subtotal(cx: Cx, code: number, refs: (V | Rng)[]): V {
  const t = Math.trunc(code)
  const c = t > 100 ? t - 100 : t
  return c < 1 || c > 11 || (t > 11 && t < 101) ? E.VALUE : over(cx, c, refs, true, false)
}

function aggregate(cx: Cx, code: number, option: number, refs: (V | Rng)[]): V {
  const c = Math.trunc(code)
  const o = Math.trunc(option)
  if (c < 1 || c > 19 || o < 0 || o > 7) return E.VALUE
  const skipNested = o < 4
  const skipErrors = o % 4 >= 2
  if (c <= 13) return over(cx, c, refs, skipNested, skipErrors)
  // LARGE … QUARTILE.EXC: one array and k, a number (an error there is #VALUE!, as in Excel)
  if (refs.length !== 2) return E.VALUE
  const kv = cx.scalar(refs[1])
  const k = isErr(kv) ? kv : toNum(kv)
  return isErr(k) ? E.VALUE : over(cx, c, [refs[0]], skipNested, skipErrors, k)
}

export const functions: Record<string, FnDef> = {
  AGGREGATE: { args: 'n n R', rep: 1, cx: true, fn: aggregate },
  SUBTOTAL: { args: 'n R', rep: 1, cx: true, fn: subtotal },
}
