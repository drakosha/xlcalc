/** The argument rule of the former Analysis ToolPak functions (bonds, EDATE, WORKDAY and the like). */
import { noBool, toDate, type Cx, type FnDef } from '../../args.js'
import type { Rules } from '../../locale/rules.js'
import { E, isErr, type CErr, type V } from '../../values.js'

const whole = (n: number | CErr) => (isErr(n) ? n : Math.trunc(n))

const SCALAR: Record<string, (v: V, rules: Rules) => unknown> = {
  v: (v) => v,
  e: (v, r) => (typeof v === 'boolean' ? E.VALUE : toDate(v, r)),
  x: noBool,
  j: (v, r) => whole(noBool(v, r)),
  n: (v, r) => r.toNum(v),
  i: (v, r) => whole(r.toNum(v)),
  d: toDate,
  b: (v, r) => r.toBool(v),
  s: (v, r) => r.toText(v),
}

/** The order in which Excel reads the arguments of a ToolPak function, where it is not the arguments' order. */
export interface AtpOrder {
  /** Read before the others, in this order (the basis of most securities, the places of the radix functions). */
  first?: readonly number[]
  /** Checks run as soon as their argument is read, given the arguments read so far: an error stops the call. */
  check?: Readonly<Record<number, (x: any, a: readonly any[]) => CErr | undefined>>
  /** `fn` gets the function context (`Cx`) first: it reads more values by the book's locale (WORKDAY's holidays). */
  cx?: boolean
}

/** A check that answers #NUM! where `ok` fails. */
export const need = (ok: (x: any, a: readonly any[]) => boolean) => (x: any, a: readonly any[]) => (ok(x, a) ? undefined : E.NUM)

/**
 * Kinds `e` date, `x` number, `j` integer reject booleans (#VALUE!, as in Excel), `v` takes any value; scalar
 * kinds are read here, in argument order unless `order` says otherwise, the first error wins; other kinds go
 * through the shared layer.
 */
export function atp(kinds: string, fn: (...a: any[]) => V, order: AtpOrder = {}): FnDef {
  const ks = kinds.split(' ').filter(Boolean)
  const local = ks.map((k) => k[0] in SCALAR)
  const first = order.first ?? []
  const seq = [...first, ...ks.map((_, i) => i).filter((i) => !first.includes(i))]
  return {
    args: ks.map((k, i) => (local[i] ? `a${k.slice(1)}` : k)).join(' '),
    cx: true,
    fn: (cx: Cx, ...a: unknown[]) => {
      for (const i of seq) {
        if (!local[i] || a[i] === undefined) continue
        const v = a[i] as V
        if (isErr(v)) return v
        const c = SCALAR[ks[i][0]](v, cx.rules)
        if (isErr(c)) return c
        a[i] = c
        const e = order.check?.[i]?.(c, a)
        if (e) return e
      }
      return order.cx ? fn(cx, ...a) : fn(...a)
    },
  }
}
