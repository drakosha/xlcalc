/** The argument rule of the former Analysis ToolPak functions (bonds, EDATE, WORKDAY and the like). */
import { coerce, noBool, toDate, type Cx, type FnDef } from '../../args.js'
import type { Rules } from '../../locale/rules.js'
import { E, isErr, Rng, type CErr, type V } from '../../values.js'

const whole = (n: number | CErr) => (isErr(n) ? n : Math.trunc(n))
const floor = (n: number | CErr) => (isErr(n) ? n : Math.floor(n))

/** Kinds this layer reads itself: `SCALAR`'s date and number kinds, and `x`, `b`, `v` by the shared `coerce`. */
const LOCAL = 'exjfbv'

const SCALAR: Record<string, (v: V, rules: Rules) => unknown> = {
  e: (v, r) => (typeof v === 'boolean' ? E.VALUE : toDate(v, r)),
  j: (v, r) => whole(noBool(v, r)),
  // rounded down: WORKDAY's days (-1E-9 is a day back)
  f: (v, r) => floor(noBool(v, r)),
}

/** The order in which Excel reads the arguments of a ToolPak function, where it is not the arguments' order. */
export interface AtpOrder {
  /** Read before the others, in this order (the basis of most securities, the places of the radix functions). */
  first?: readonly number[]
  /** Checks run as soon as their argument is read, given the arguments read so far: an error stops the call. */
  check?: Readonly<Record<number, (x: any, a: readonly any[]) => CErr | undefined>>
  /** `fn` gets the function context (`Cx`) first: it reads more values by the book's locale (WORKDAY's holidays). */
  cx?: boolean
  /** The answer is a date (`FnDef.date`). */
  date?: FnDef['date']
}

/** A check that answers #NUM! where `ok` fails. */
export const need = (ok: (x: any, a: readonly any[]) => boolean) => (x: any, a: readonly any[]) => (ok(x, a) ? undefined : E.NUM)

/** Kind `e`, a date: it gets its argument as given, a reference too (`r`), to tell date text in a cell from date text given as a value. */
const DATES = 'e'

/**
 * A read argument as the function takes it. In a 1904 book, date text given as a value (typed in the formula or
 * computed, not read from a cell) is taken as the serial number it reads in the book's system without moving it into
 * the 1900 calendar, as Excel does (the workbook set `date1904-text`): EDATE("1/31/2025",1) is 42793 there,
 * EDATE(A1,1) with that text in A1 is 44254.
 */
function taken(kind: string, x: unknown, v: V, cx: Cx): unknown {
  const c = kind in SCALAR ? SCALAR[kind](v, cx.rules) : coerce(kind, v, cx)
  return DATES.includes(kind) && typeof x === 'string' && !isErr(c) ? (c as number) - cx.rules.epoch : c
}

/**
 * Kinds `e` date, `j` integer (`f` rounded down) reject booleans (#VALUE!, as in Excel), as does `x` number;
 * `b` boolean and `v` any value go through the shared `coerce`. Scalar kinds are read here, in argument order
 * unless `order` says otherwise, the first error wins; other kinds go through the shared layer.
 */
export function atp(kinds: string, fn: (...a: any[]) => V, order: AtpOrder = {}): FnDef {
  const ks = kinds.split(' ').filter(Boolean)
  const local = ks.map((k) => LOCAL.includes(k[0]))
  const first = order.first ?? []
  const seq = [...first, ...ks.map((_, i) => i).filter((i) => !first.includes(i))]
  return {
    args: ks.map((k, i) => (local[i] ? `${DATES.includes(k[0]) ? 'r' : 'a'}${k.slice(1)}` : k)).join(' '),
    cx: true,
    ...(order.date ? { date: order.date } : {}),
    fn: (cx: Cx, ...a: unknown[]) => {
      for (const i of seq) {
        if (!local[i] || a[i] === undefined) continue
        // a date kind gets a reference as it is: a range of several cells is #VALUE!, as a value
        const x = a[i]
        const v = (x instanceof Rng ? cx.scalar(x) : x) as V
        if (isErr(v)) return v
        const c = taken(ks[i][0], x, v, cx)
        if (isErr(c)) return c
        a[i] = c
        const e = order.check?.[i]?.(c, a)
        if (e) return e
      }
      return order.cx ? fn(cx, ...a) : fn(...a)
    },
  }
}
