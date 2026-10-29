/**
 * Sums and products. SUM and PRODUCT ported from formula.js 4.6.1 (src/math-trig.js), Copyright (c) 2014
 * Sutoiku, Inc., MIT License (THIRD_PARTY.md); the rest this engine's own.
 */
import { occupied, oneCell, type Arg, type Cx, type FnDef, type Grid, type Segment } from '../../args.js'
import { Arr, asArray } from '../../arrays/value.js'
import { spend, STRIDE } from '../../exec/budget.js'
import type { Rules } from '../../locale/rules.js'
import { E, isErr, type CErr, type V } from '../../values.js'
import { excelSum, product, sumSquares } from '../common.js'
import { aggregateIf, aggregateIfs } from '../criteria.js'
import { numberPairs } from '../statistical/common.js'

/** The sum of the numbers at the selected positions; an error there is the answer. */
function sumAt(values: readonly V[], sel: readonly number[]): V {
  let acc = 0
  for (const p of sel) {
    const v = values[p]
    if (typeof v === 'number') acc += v
    else if (isErr(v)) return v
  }
  return acc
}

/**
 * SUMPRODUCT: arrays of one shape; anything but a number counts as 0. Each array in turn answers a blank cell given
 * alone with #VALUE! and its first error, before the shapes are compared (the sets `sumx-cell-held`, `sumproduct-held`:
 * `SUMPRODUCT(A1:A3,B1:B2)` with an error in A2 is that error). Past the occupied parts every array holds its fill
 * (blank for a range), so `SUMPRODUCT((A:A="x")*B:B)` computes the occupied rows only. Long products are a long
 * loop: the budget's deadline and signal are checked inside it.
 */
function sumproduct(cx: Cx, xs: Arg[]): V {
  const arrs: Arr[] = []
  for (const x of xs) {
    if (isErr(x)) return x
    const a = asArray(cx.view(x))
    const e = oneCell(x, cx) ?? a.v.find(isErr) ?? (a.h * a.w < a.rows * a.cols && isErr(a.fill) ? a.fill : undefined)
    if (e) return e as CErr
    arrs.push(a)
  }
  const [first] = arrs
  if (arrs.some((a) => a.rows !== first.rows || a.cols !== first.cols)) return E.VALUE
  const num = (v: V) => (typeof v === 'number' ? v : 0)
  const end = Math.max(...arrs.map(occupied))
  let acc = 0
  if (arrs.every((a) => a.w === a.cols || a.h === 0)) {
    // each stored in its whole width: position i is v[i] in the stored rows, the fill past them
    const stored = arrs.map((a) => a.h * a.w)
    for (let i = 0; i < end; i++) {
      if ((i & STRIDE) === 0) spend(Math.min(STRIDE + 1, end - i))
      let p = 1
      for (let k = 0; k < arrs.length; k++) p *= num(i < stored[k] ? arrs[k].v[i] : arrs[k].fill)
      acc += p
    }
  } else
    for (let i = 0; i < end; i++) {
      if ((i & STRIDE) === 0) spend(Math.min(STRIDE + 1, end - i))
      const r = Math.floor(i / first.cols)
      let p = 1
      for (const a of arrs) p *= num(a.at(r, i - r * first.cols))
      acc += p
    }
  const rest = first.rows * first.cols - end
  return rest > 0 ? acc + rest * arrs.reduce((p, a) => p * num(a.fill), 1) : acc
}

/** SUMX2MY2 and its kin: pairs of numbers (rule 4) as CORREL takes them; no pairs at all is #DIV/0!. */
const sumx = (f: (x: number, y: number) => number): FnDef => ({
  args: 'C C',
  fn: (a: Grid, b: Grid) => {
    const p = numberPairs(a, b)
    if (isErr(p)) return p
    const [xs, ys] = p
    return xs.length ? xs.reduce((s, x, i) => s + f(x, ys[i]), 0) : E.DIV0
  },
})

/** x^p for SERIESSUM: 0 to a power not above 0 is #NUM!, a negative x to a fractional power adds 0 (Excel). */
function term(x: number, p: number): number | CErr {
  if (x === 0) return p > 0 ? 0 : E.NUM
  return x < 0 && !Number.isInteger(p) ? 0 : Math.pow(x, p)
}

function seriessum(cx: Cx, x: number, n: number, m: number, coefs: Segment): V {
  // a range gives its numbers in order, blanks left out; text or a boolean there is #VALUE!
  const cs: number[] = []
  if (Array.isArray(coefs)) {
    for (const v of coefs as readonly V[]) {
      if (isErr(v)) return v
      if (typeof v === 'number') cs.push(v)
      else if (v !== null) return E.VALUE
    }
  } else {
    const c = typeof coefs === 'boolean' ? E.VALUE : cx.rules.toNum(coefs as V)
    if (isErr(c)) return c
    cs.push(c)
  }
  let acc = 0
  for (let i = 0; i < cs.length; i++) {
    const t = term(x, n + i * m)
    if (isErr(t)) return t
    acc += cs[i] * t
  }
  return acc
}

/** PERCENTOF's sum: a range gives its numbers, a direct argument is coerced; an error in a range or array is #NUM!, a direct one itself (Excel). */
function percentSum(s: Segment, rules: Rules): number | CErr {
  if (!Array.isArray(s)) return rules.toNum(s as V)
  let acc = 0
  for (const v of s as readonly V[]) {
    if (isErr(v)) return E.NUM
    if (typeof v === 'number') acc += v
  }
  return acc
}

export const functions: Record<string, FnDef> = {
  PERCENTOF: {
    args: 'V V',
    cx: true,
    fn: (cx: Cx, part: Segment, all: Segment) => {
      const a = percentSum(part, cx.rules)
      if (isErr(a)) return a
      const b = percentSum(all, cx.rules)
      if (isErr(b)) return b
      return b === 0 ? E.DIV0 : a / b
    },
  },
  PRODUCT: { args: 'N', rep: 1, fn: product },
  SERIESSUM: { args: 'x x x V', cx: true, fn: seriessum },
  SUM: { args: 'N', rep: 1, fn: excelSum },
  SUMIF: {
    args: 'R a R?',
    cx: true,
    criteria: true,
    fn: aggregateIf(sumAt),
  },
  SUMIFS: {
    args: 'R R a',
    rep: 2,
    cx: true,
    criteria: true,
    fn: aggregateIfs(sumAt),
  },
  SUMPRODUCT: { args: 'R', rep: 1, cx: true, fn: sumproduct },
  SUMSQ: { args: 'N', rep: 1, fn: sumSquares },
  SUMX2MY2: sumx((x, y) => x * x - y * y),
  SUMX2PY2: sumx((x, y) => x * x + y * y),
  SUMXMY2: sumx((x, y) => (x - y) * (x - y)),
}
