/** Factorials, binomial coefficients, GCD and LCM with Excel's argument checks. This engine's own. */
import { gridAt, type Arg, type Cx, type FnDef } from '../../args.js'
import { E, isErr, toNum, whole, type CErr, type V } from '../../values.js'
import { fact } from '../common.js'

/** Excel: COMBIN and COMBINA take n below 2^31 − 1. */
const MAX_N = 2 ** 31 - 1
/** Excel: GCD and LCM take values up to 2^53, and LCM's answer as well. */
const MAX_INT = 2 ** 53
const MAX = Number.MAX_VALUE

/** C(n, k) by a product of quotients, dividing first where a product would overflow (C(1029, 514) is finite). */
function choose(n: number, k: number): number {
  const kk = Math.min(k, n - k)
  let r = 1
  for (let i = 1; i <= kk && r !== Infinity; i++) {
    const f = n - kk + i
    r = r > MAX / f ? (r / i) * f : (r * f) / i
  }
  return r
}

function combin(n: number, k: number): number | CErr {
  const a = Math.trunc(n)
  const b = Math.trunc(k)
  return a < 0 || b < 0 || a < b || a >= MAX_N ? E.NUM : choose(a, b)
}

function combina(n: number, k: number): number | CErr {
  const a = Math.trunc(n)
  const b = Math.trunc(k)
  if (a < 0 || b < 0 || (a === 0 && b > 0)) return E.NUM
  return a === 0 ? 1 : combin(a + b - 1, b)
}

function factdouble(x: number): number | CErr {
  if (x < -1) return E.NUM
  let p = 1
  for (let k = Math.trunc(x); k > 1 && p !== Infinity; k -= 2) p *= k
  return p
}

/**
 * The whole numbers of GCD, LCM and MULTINOMIAL: numbers and numeric text, in ranges too; a blank cell of a range
 * is 0, a blank single cell is left out (Excel: LCM(A1:A3) with A2 blank is 0, LCM(blank,1) is 1); an error, a
 * boolean or other text stops in the order Excel reads them, values inside ranges and arrays too
 * (`GCD({4,"a"},1/0)` is #VALUE!), and only then a negative number is #NUM!.
 */
function wholes(cx: Cx, args: readonly Arg[]): number[] | CErr {
  const out: number[] = []
  let negative = false
  for (const x of args) {
    if (isErr(x)) return x
    const g = cx.grid(x)
    const n = g.rows * g.cols
    if (n === 1 && gridAt(g, 0) === null) continue
    const end = Math.min(n, g.h === 0 ? 0 : (g.h - 1) * g.cols + g.w)
    for (let i = 0; i < end; i++) {
      const v: V = gridAt(g, i)
      const x = typeof v === 'boolean' ? E.VALUE : toNum(v)
      if (isErr(x)) return x
      negative ||= x < 0
      out.push(Math.trunc(x))
    }
    if (n > end) out.push(0)
  }
  return negative ? E.NUM : out
}

function gcd2(a: number, b: number): number {
  while (b) [a, b] = [b, a % b]
  return a
}

export const functions: Record<string, FnDef> = {
  COMBIN: { args: 'n n', fn: combin },
  COMBINA: { args: 'n n', fn: combina },
  FACT: { args: 'n', fn: (x: number) => (x < 0 ? E.NUM : fact(whole(x))) },
  FACTDOUBLE: { args: 'x', fn: factdouble },
  GCD: {
    args: 'R',
    rep: 1,
    cx: true,
    fn: (cx: Cx, args: Arg[]) => {
      const xs = wholes(cx, args)
      if (isErr(xs)) return xs
      return xs.some((x) => x > MAX_INT) ? E.NUM : xs.reduce(gcd2, 0)
    },
  },
  LCM: {
    args: 'R',
    rep: 1,
    cx: true,
    fn: (cx: Cx, args: Arg[]) => {
      const xs = wholes(cx, args)
      if (isErr(xs)) return xs
      let l = 1
      for (const x of xs) {
        if (x > MAX_INT) return E.NUM
        l = x === 0 || l === 0 ? 0 : (l / gcd2(l, x)) * x
      }
      return l > MAX_INT ? E.NUM : l
    },
  },
  MULTINOMIAL: {
    args: 'R',
    rep: 1,
    cx: true,
    fn: (cx: Cx, args: Arg[]) => {
      const xs = wholes(cx, args)
      if (isErr(xs)) return xs
      const total = xs.reduce((s, x) => s + x, 0)
      return fact(total) / xs.reduce((p, x) => p * fact(x), 1)
    },
  },
}
