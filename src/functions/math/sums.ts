/**
 * Sums and products. SUM and PRODUCT ported from formula.js 4.6.1 (src/math-trig.js), Copyright (c) 2014
 * Sutoiku, Inc., MIT License (THIRD_PARTY.md); the rest this engine's own.
 */
import { gridAt, pairs, type Cx, type FnDef, type Grid, type Segment } from '../../args.js'
import { MAX_COLS, MAX_ROWS } from '../../model.js'
import { E, isErr, Rng, toNum, type CErr, type V } from '../../values.js'

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

const asRng = (x: V | Rng): Rng | CErr => (x instanceof Rng ? x : isErr(x) ? x : E.VALUE)

/** Positions of a grid that hold cells, row by row in the referenced width; past them everything is blank. */
const filled = (g: Grid) => (g.h === 0 ? 0 : (g.h - 1) * g.cols + g.w)

/** The first error of the grids, array by array (Excel: SUMPRODUCT(A,B) shows A's error before B's). */
function firstError(grids: readonly Grid[]): CErr | undefined {
  for (const g of grids) {
    const e = g.v.find(isErr)
    if (e) return e as CErr
  }
}

/** SUMPRODUCT: arrays of one shape; anything but a number counts as 0. */
function sumproduct(grids: Grid[]): V {
  const [first] = grids
  if (grids.some((g) => g.rows !== first.rows || g.cols !== first.cols)) return E.VALUE
  const e = firstError(grids)
  if (e) return e
  const end = Math.max(...grids.map(filled))
  let acc = 0
  for (let i = 0; i < end; i++) {
    let p = 1
    for (const g of grids) {
      const v = gridAt(g, i)
      p *= typeof v === 'number' ? v : 0
    }
    acc += p
  }
  return acc
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

function seriessum(x: number, n: number, m: number, coefs: Segment): V {
  // a range gives its numbers in order, blanks left out; text or a boolean there is #VALUE!
  const cs: number[] = []
  if (Array.isArray(coefs)) {
    for (const v of coefs as readonly V[]) {
      if (isErr(v)) return v
      if (typeof v === 'number') cs.push(v)
      else if (v !== null) return E.VALUE
    }
  } else {
    const c = typeof coefs === 'boolean' ? E.VALUE : toNum(coefs as V)
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

/** PERCENTOF's sum: a range gives its numbers, a direct argument is coerced; any error is #NUM! (Excel). */
function percentSum(s: Segment): number | CErr {
  if (!Array.isArray(s)) return isErr(s) ? E.NUM : toNum(s as V)
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
    fn: (part: Segment, all: Segment) => {
      const a = percentSum(part)
      if (isErr(a)) return a
      const b = percentSum(all)
      if (isErr(b)) return b
      return b === 0 ? E.DIV0 : a / b
    },
  },
  PRODUCT: { args: 'N', rep: 1, fn: (xs: number[]) => (xs.length ? xs.reduce((p, x) => p * x, 1) : 0) },
  SERIESSUM: { args: 'x x x V', fn: seriessum },
  SUM: { args: 'N', rep: 1, fn: (xs: number[]) => xs.reduce((s, x) => s + x, 0) },
  // the sum range takes the size of the criteria range from its own top-left cell, as in Excel
  SUMIF: {
    args: 'R a R?',
    cx: true,
    fn: (cx: Cx, range: V | Rng, crit: V | Rng, sum?: V | Rng) => {
      const cg = asRng(range)
      if (isErr(cg)) return cg
      if (sum === undefined) {
        const r = cx.select([[cg, crit]])
        return isErr(r) ? r : sumAt(r.read(cg), r.sel)
      }
      const sg = asRng(sum)
      if (isErr(sg)) return sg
      // past the sheet's edge the sum range is cut, and then it no longer fits: #VALUE!
      const sized = new Rng(sg.s, sg.r1, sg.c1, Math.min(sg.r1 + cg.r2 - cg.r1, MAX_ROWS - 1), Math.min(sg.c1 + cg.c2 - cg.c1, MAX_COLS - 1))
      const r = cx.select([[cg, crit]], sized)
      return isErr(r) ? r : sumAt(r.read(sized), r.sel)
    },
  },
  SUMIFS: {
    args: 'R R a',
    rep: 2,
    cx: true,
    fn: (cx: Cx, sum: V | Rng, rest: (V | Rng)[]) => {
      const sg = asRng(sum)
      if (isErr(sg)) return sg
      const pairs: [V | Rng, V | Rng][] = []
      for (let i = 0; i < rest.length; i += 2) pairs.push([rest[i], rest[i + 1]])
      const r = cx.select(pairs, sg)
      return isErr(r) ? r : sumAt(r.read(sg), r.sel)
    },
  },
  SUMPRODUCT: { args: 'G', rep: 1, fn: sumproduct },
  SUMSQ: { args: 'N', rep: 1, fn: (xs: number[]) => xs.reduce((s, x) => s + x * x, 0) },
  SUMX2MY2: sumx((x, y) => x * x - y * y),
  SUMX2PY2: sumx((x, y) => x * x + y * y),
  SUMXMY2: sumx((x, y) => (x - y) * (x - y)),
}
