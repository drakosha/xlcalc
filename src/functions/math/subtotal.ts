/**
 * SUBTOTAL and AGGREGATE on the statistical functions. A cell whose formula calls either of them (anywhere in the
 * formula) is skipped as nested, as in Excel. Hidden rows (the workbook set `hidden`): SUBTOTAL 101–111 and
 * AGGREGATE's options 1, 3, 5, 7 skip them all, SUBTOTAL 1–11 only on a filtered sheet, AGGREGATE's even options
 * never; hidden columns count. AGGREGATE 14–19 take an array too, which has no rows to skip. This engine's own.
 */
import { Arr } from '../../arrays/value.js'
import { E, isErr, Rng, whole, type CErr, type V } from '../../values.js'
import type { Cx, FnDef, Grid, HiddenRule } from '../../args.js'
import { NEWER_MAX_ARGS } from '../common.js'
import { functions as stat } from '../statistical/descriptive.js'
import { functions as sums } from './sums.js'

/** A statistical function applied to the numbers of the references, with its rule for no numbers. */
function apply(name: string, xs: number[], k?: number): V {
  const def = name === 'SUM' || name === 'PRODUCT' ? sums[name] : stat[name]
  if (def.empty && !xs.length) return E[def.empty]
  // the statistical functions answer one value
  if (name === 'MODE.SNGL') return def.fn([{ rows: 1, cols: xs.length, h: 1, w: xs.length, v: xs } satisfies Grid]) as V
  return (k === undefined ? def.fn(xs) : def.fn(xs, k)) as V
}

const NAMES = ['AVERAGE', 'COUNT', 'COUNTA', 'MAX', 'MIN', 'PRODUCT', 'STDEV.S', 'STDEV.P', 'SUM', 'VAR.S', 'VAR.P', 'MEDIAN', 'MODE.SNGL', 'LARGE', 'SMALL', 'PERCENTILE.INC', 'QUARTILE.INC', 'PERCENTILE.EXC', 'QUARTILE.EXC']

/** What SUBTOTAL and AGGREGATE skip in their references. */
interface Skip {
  nested: boolean
  hidden: HiddenRule
  errors: boolean
}

/**
 * The function `code` (1–13) over the references' values: COUNT counts numbers, COUNTA every non-blank cell
 * (errors too), the rest take the numbers and stop at the first error, unless errors are skipped. AGGREGATE 14–19
 * (`k`) also take an array.
 */
function over(cx: Cx, code: number, refs: readonly (V | Rng | Arr)[], skip: Skip, k?: number): V {
  const xs: number[] = []
  let filled = 0
  let first: CErr | undefined
  for (const x of refs) {
    // only references, and an array for 14–19: a value in their place is #VALUE!
    if (!(x instanceof Rng) && !(x instanceof Arr && k !== undefined)) return isErr(x) ? x : E.VALUE
    for (const v of x instanceof Rng ? cx.counted(x, skip.nested, skip.hidden) : cx.grid(x).v) {
      if (v === null) continue
      if (isErr(v)) {
        if (skip.errors) continue
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
  const t = whole(code)
  const c = t > 100 ? t - 100 : t
  return c < 1 || c > 11 || (t > 11 && t < 101) ? E.VALUE : over(cx, c, refs, { nested: true, hidden: t > 100 ? 'all' : 'filtered', errors: false })
}

function aggregate(cx: Cx, code: number, option: number, refs: (V | Rng | Arr)[]): V {
  const c = whole(code)
  const o = Math.trunc(option)
  if (c < 1 || c > 19 || o < 0 || o > 7) return E.VALUE
  const skip: Skip = { nested: o < 4, hidden: o % 2 === 1 ? 'all' : 'none', errors: o % 4 >= 2 }
  if (c <= 13) return over(cx, c, refs, skip)
  // LARGE … QUARTILE.EXC: one array and k, a number (an error there is #VALUE!, as in Excel); a range as k
  // gives its top-left cell (Excel 365: AGGREGATE(14,6,M1:M10,A1:A3) is LARGE with k from A1)
  if (refs.length !== 2) return E.VALUE
  const r = refs[1]
  const kv = cx.scalar(r instanceof Rng ? new Rng(r.s, r.r1, r.c1, r.r1, r.c1) : r)
  const k = isErr(kv) ? kv : cx.rules.toNum(kv)
  return isErr(k) ? E.VALUE : over(cx, c, [refs[0]], skip, k)
}

export const functions: Record<string, FnDef> = {
  AGGREGATE: { args: 'n n R', rep: 1, max: NEWER_MAX_ARGS, cx: true, fn: aggregate },
  SUBTOTAL: { args: 'n R', rep: 1, cx: true, fn: subtotal },
}
