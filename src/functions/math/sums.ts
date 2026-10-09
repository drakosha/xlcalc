/**
 * Sums and products. SUM and PRODUCT ported from formula.js 4.6.1 (src/math-trig.js), Copyright (c) 2014
 * Sutoiku, Inc., MIT License (THIRD_PARTY.md); the rest this engine's own.
 */
import { pairs, type Arg, type Cx, type FnDef, type Grid, type Segment } from '../../args.js'
import { Arr, asArray } from '../../arrays/value.js'
import type { Rules } from '../../locale/rules.js'
import { E, isErr, type CErr, type V } from '../../values.js'
import { product, sum, sumSquares } from '../common.js'
import { aggregateIf, aggregateIfs } from '../criteria.js'

/** The sum of the numbers at the selected positions; an error there is the answer. */
function sumAt(values: readonly V[], sel: number[]): V {
  let acc = 0
  for (const p of sel) {
    const v = values[p]
    if (typeof v === 'number') acc += v
    else if (isErr(v)) return v
  }
  return acc
}

/** Positions of a grid that hold cells, row by row in the referenced width; past them everything is blank. */
const filled = (g: Grid) => (g.h === 0 ? 0 : (g.h - 1) * g.cols + g.w)

/** The first error of the grids, array by array (Excel: SUMPRODUCT(A,B) shows A's error before B's). */
function firstError(grids: readonly Grid[]): CErr | undefined {
  for (const g of grids) {
    const e = g.v.find(isErr)
    if (e) return e as CErr
  }
}

/**
 * SUMPRODUCT: arrays of one shape; anything but a number counts as 0. Past the occupied parts every array
 * holds its fill (blank for a range), so `SUMPRODUCT((A:A="x")*B:B)` computes the occupied rows only.
 */
function sumproduct(cx: Cx, xs: Arg[]): V {
  const arrs: Arr[] = []
  for (const x of xs) {
    if (isErr(x)) return x
    arrs.push(asArray(cx.view(x)))
  }
  const [first] = arrs
  if (arrs.some((a) => a.rows !== first.rows || a.cols !== first.cols)) return E.VALUE
  for (const a of arrs) {
    const e = a.v.find(isErr) ?? (a.h * a.w < a.rows * a.cols && isErr(a.fill) ? a.fill : undefined)
    if (e) return e as CErr
  }
  const num = (v: V) => (typeof v === 'number' ? v : 0)
  const end = Math.max(...arrs.map(filled))
  let acc = 0
  for (let i = 0; i < end; i++) {
    const r = Math.floor(i / first.cols)
    let p = 1
    for (const a of arrs) p *= num(a.at(r, i - r * first.cols))
    acc += p
  }
  const rest = first.rows * first.cols - end
  return rest > 0 ? acc + rest * arrs.reduce((p, a) => p * num(a.fill), 1) : acc
}

/** SUMX2MY2 and its kin: pairs of numbers (rule 4); no pairs at all is #DIV/0!. */
const sumx = (f: (x: number, y: number) => number): FnDef => ({
  args: 'G G',
  fn: (a: Grid, b: Grid) => {
    if (a.rows * a.cols !== b.rows * b.cols) return E.NA
    const e = firstError([a, b])
    if (e) return e
    const p = pairs(a, b)
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
  SUM: { args: 'N', rep: 1, fn: sum },
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
