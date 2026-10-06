/**
 * Correlation, covariance and the simple linear regression, after formula.js 4.6.1
 * (https://github.com/formulajs/formulajs), src/statistical.js: Copyright (c) 2014 Sutoiku, Inc., MIT License
 * (THIRD_PARTY.md). Pairs and errors follow Excel: an error in a range is the answer, pairs with a non-number
 * are dropped, ranges of different sizes are #N/A.
 */
import { codev, devsq, mean, numberPairs } from './common.js'
import type { FnDef, Grid } from '../../args.js'
import { E, isErr, type CErr, type V } from '../../values.js'

interface Fit {
  n: number
  mx: number
  my: number
  sxx: number
  syy: number
  sxy: number
}

/** Sums of the simple regression of y on x over the number pairs. */
function fit(gy: Grid, gx: Grid): Fit | CErr {
  const p = numberPairs(gy, gx)
  if (isErr(p)) return p
  const [ys, xs] = p
  if (!xs.length) return E.DIV0
  return { n: xs.length, mx: mean(xs), my: mean(ys), sxx: devsq(xs), syy: devsq(ys), sxy: codev(xs, ys) }
}

const on = (g: (f: Fit) => V) => (a: Grid, b: Grid) => {
  const f = fit(a, b)
  return isErr(f) ? f : g(f)
}

const correl = on((f) => (f.sxx === 0 || f.syy === 0 ? E.DIV0 : f.sxy / Math.sqrt(f.sxx * f.syy)))
const slope = (f: Fit) => (f.sxx === 0 ? E.DIV0 : f.sxy / f.sxx)

/** FORECAST(x, known_y, known_x). */
function forecast(x: number, gy: Grid, gx: Grid): V {
  const f = fit(gy, gx)
  if (isErr(f)) return f
  return f.sxx === 0 ? E.DIV0 : f.my + (f.sxy / f.sxx) * (x - f.mx)
}

export const functions: Record<string, FnDef> = {
  CORREL: { args: 'G G', fn: correl },
  'COVARIANCE.P': { args: 'G G', fn: on((f) => f.sxy / f.n) },
  'COVARIANCE.S': { args: 'G G', fn: on((f) => (f.n < 2 ? E.DIV0 : f.sxy / (f.n - 1))) },
  FORECAST: { args: 'n G G', fn: forecast },
  'FORECAST.LINEAR': { args: 'n G G', fn: forecast },
  INTERCEPT: {
    args: 'G G',
    fn: on((f) => {
      const b = slope(f)
      return isErr(b) ? b : f.my - b * f.mx
    }),
  },
  PEARSON: { args: 'G G', fn: correl },
  RSQ: {
    args: 'G G',
    fn: (a: Grid, b: Grid) => {
      const r = correl(a, b)
      return typeof r === 'number' ? r * r : r
    },
  },
  SLOPE: { args: 'G G', fn: on(slope) },
  STEYX: { args: 'G G', fn: on((f) => (f.n < 3 || f.sxx === 0 ? E.DIV0 : Math.sqrt(Math.max(0, f.syy - (f.sxy * f.sxy) / f.sxx) / (f.n - 2)))) },
}
