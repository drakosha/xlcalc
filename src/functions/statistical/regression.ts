/**
 * Correlation, covariance and the simple linear regression, after formula.js 4.6.1
 * (https://github.com/formulajs/formulajs), src/statistical.js: Copyright (c) 2014 Sutoiku, Inc., MIT License
 * (THIRD_PARTY.md). Pairs and errors follow Excel: an error in a range is the answer, pairs with a non-number
 * are dropped, ranges of different sizes are #N/A.
 */
import { numberPairs } from './common.js'
import { correlation, covariance, pairMoments, prediction, residualDeviation, slopeOf, type PairMoments } from './moments.js'
import type { FnDef, Grid } from '../../args.js'
import { E, isErr, type CErr, type V } from '../../values.js'

type Fit = PairMoments

/** Sums of the simple regression of y on x over the number pairs. */
function fit(gy: Grid, gx: Grid): Fit | CErr {
  const p = numberPairs(gy, gx)
  if (isErr(p)) return p
  const [ys, xs] = p
  if (!xs.length) return E.DIV0
  return pairMoments(xs, ys)
}

const on = (g: (f: Fit) => V) => (a: Grid, b: Grid) => {
  const f = fit(a, b)
  return isErr(f) ? f : g(f)
}

const correl = on((f) => correlation(f) ?? E.DIV0)
const slope = (f: Fit) => slopeOf(f) ?? E.DIV0

/** FORECAST(x, known_y, known_x). */
function forecast(x: number, gy: Grid, gx: Grid): V {
  const f = fit(gy, gx)
  if (isErr(f)) return f
  return prediction(f, x) ?? E.DIV0
}

export const functions: Record<string, FnDef> = {
  CORREL: { args: 'G G', fn: correl },
  'COVARIANCE.P': { args: 'G G', fn: on((f) => covariance(f, f.n)) },
  'COVARIANCE.S': { args: 'G G', fn: on((f) => (f.n < 2 ? E.DIV0 : covariance(f, f.n - 1))) },
  FORECAST: { args: 'n G G', fn: forecast },
  'FORECAST.LINEAR': { args: 'n G G', fn: forecast },
  INTERCEPT: {
    args: 'G G',
    fn: on((f) => prediction(f, 0) ?? E.DIV0),
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
  STEYX: { args: 'G G', fn: on((f) => (f.n < 3 ? E.DIV0 : residualDeviation(f) ?? E.DIV0)) },
}
