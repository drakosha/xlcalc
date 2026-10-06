/**
 * Distribution functions (densities, mass functions and their cumulative forms). BINOM.DIST.RANGE, GAUSS,
 * GAMMALN, PHI and PROB, and the parameter names of the rest, follow formula.js 4.6.1
 * (https://github.com/formulajs/formulajs), src/statistical.js: Copyright (c) 2014 Sutoiku, Inc., MIT License
 * (THIRD_PARTY.md); the numerics are `./special.ts`, the domains Excel's.
 */
import { gammaln, gammafn } from '../jstat.js'
import { gridAt, type FnDef, type Grid } from '../../args.js'
import { E, type V } from '../../values.js'
import { combin } from '../common.js'
import { betaPrefix, dbinom, dgamma, dpois, gammaPQ, ibeta, lbeta, normCdf, normPdf } from './special.js'

/** Excel's ceiling on degrees of freedom of the chi-squared, t and F distributions. */
export const MAX_DF = 1e10
/** Counts at or past 2^31 − 1 are #NUM! in BINOM.DIST, PERMUT and PERMUTATIONA. */
export const MAX_COUNT = 2147483647

export function normDist(x: number, m: number, sd: number, cumulative: boolean): V {
  if (sd <= 0) return E.NUM
  const z = (x - m) / sd
  if (cumulative) return normCdf(z)[0]
  return Number.isFinite(z * z) ? normPdf(z) / sd : E.NUM
}

export const normSDist = (z: number, cumulative: boolean) => normDist(z, 0, 1, cumulative)

export function lognormDist(x: number, m: number, sd: number, cumulative: boolean): V {
  if (x <= 0 || sd <= 0) return E.NUM
  const z = (Math.log(x) - m) / sd
  return cumulative ? normCdf(z)[0] : normPdf(z) / (x * sd)
}

export function exponDist(x: number, rate: number, cumulative: boolean): V {
  if (x < 0 || rate <= 0) return E.NUM
  const t = rate * x
  if (!Number.isFinite(t)) return E.NUM
  return cumulative ? -Math.expm1(-t) : rate * Math.exp(-t)
}

/** WEIBULL.DIST as Excel computes it: the density is α x^(α−1)/β^α e^(−t), #DIV/0! when β^α vanishes, #NUM! when a power overflows, 0 at 0. */
export function weibullDist(x: number, alpha: number, beta: number, cumulative: boolean): V {
  if (x < 0 || alpha <= 0 || beta <= 0) return E.NUM
  const t = (x / beta) ** alpha
  if (cumulative) return Number.isFinite(t) ? -Math.expm1(-t) : E.NUM
  if (x === 0) return 0
  const ba = beta ** alpha
  if (ba === 0) return E.DIV0
  const xa = x ** (alpha - 1)
  return Number.isFinite(ba) && Number.isFinite(xa) && Number.isFinite(t) ? ((alpha * xa) / ba) * Math.exp(-t) : E.NUM
}

export function gammaDist(x: number, a: number, b: number, cumulative: boolean): V {
  if (x < 0 || a <= 0 || b <= 0) return E.NUM
  if (cumulative) return gammaPQ(a, x / b)[0]
  if (x === 0) return a <= 1 ? E.NUM : 0
  return dgamma(x, a, b)
}

/** Degrees of freedom: truncated, from 1 to 10^10. */
export const badDf = (df: number) => df < 1 || df > MAX_DF

export function chisqDist(x: number, df: number, cumulative: boolean): V {
  if (x < 0 || badDf(df)) return E.NUM
  if (cumulative) return gammaPQ(df / 2, x / 2)[0]
  // the density at 0: infinite for 1 degree of freedom, ½ for 2
  if (x === 0) return df === 1 ? E.NUM : df === 2 ? 0.5 : 0
  return dgamma(x, df / 2, 2)
}

export const chisqDistRt = (x: number, df: number): V => (x < 0 || badDf(df) ? E.NUM : gammaPQ(df / 2, x / 2)[1])

/** [P(T ≤ t), P(T > t)] for Student's t. */
export function tCdf(t: number, df: number): [number, number] {
  const t2 = t * t
  const tail = 0.5 * ibeta(df / (df + t2), t2 / (df + t2), df / 2, 0.5)[0]
  return t < 0 ? [tail, 1 - tail] : [1 - tail, tail]
}

export function tPdf(t: number, df: number): number {
  const t2 = t * t
  if (t === 0) return Math.exp(-lbeta(df / 2, 0.5)) / Math.sqrt(df)
  return betaPrefix(df / (df + t2), t2 / (df + t2), df / 2, 0.5) / Math.abs(t)
}

export function tDist(t: number, df: number, cumulative: boolean): V {
  if (!Number.isFinite(t * t)) return E.NUM
  if (cumulative) return badDf(df) ? E.NUM : tCdf(t, df)[0]
  // Excel divides by the degrees of freedom first: 0 is #DIV/0!
  if (df === 0) return E.DIV0
  return df < 1 ? E.NUM : tPdf(t, df)
}

export const tDistRt = (t: number, df: number): V => (!Number.isFinite(t * t) || badDf(df) ? E.NUM : tCdf(t, df)[1])
export const tDist2T = (t: number, df: number): V => (t < 0 || !Number.isFinite(t * t) || badDf(df) ? E.NUM : 2 * tCdf(t, df)[1])

/** [P(F ≤ x), P(F > x)] for the F distribution. */
export const fCdf = (x: number, d1: number, d2: number) => ibeta((d1 * x) / (d1 * x + d2), d2 / (d1 * x + d2), d1 / 2, d2 / 2)

export const fPdf = (x: number, d1: number, d2: number) => betaPrefix((d1 * x) / (d1 * x + d2), d2 / (d1 * x + d2), d1 / 2, d2 / 2) / x

export function fDist(x: number, d1: number, d2: number, cumulative: boolean): V {
  if (x < 0 || badDf(d1) || badDf(d2)) return E.NUM
  if (cumulative) return fCdf(x, d1, d2)[0]
  if (x === 0) return d1 === 1 ? E.NUM : d1 === 2 ? 1 : 0
  return fPdf(x, d1, d2)
}

export const fDistRt = (x: number, d1: number, d2: number): V => (x < 0 || badDf(d1) || badDf(d2) ? E.NUM : fCdf(x, d1, d2)[1])

export function betaDist(x: number, a: number, b: number, cumulative: boolean, lo = 0, hi = 1): V {
  if (a <= 0 || b <= 0 || lo >= hi || x < lo || x > hi) return E.NUM
  const z = (x - lo) / (hi - lo)
  const y = (hi - x) / (hi - lo)
  if (cumulative) return ibeta(z, y, a, b)[0]
  if (z === 0) return a <= 1 ? E.NUM : 0
  if (y === 0) return b <= 1 ? E.NUM : 0
  return betaPrefix(z, y, a, b) / (z * y) / (hi - lo)
}

export function binomDist(k: number, n: number, p: number, cumulative: boolean): V {
  if (k < 0 || k > n || n >= MAX_COUNT || p < 0 || p > 1) return E.NUM
  if (!cumulative) return dbinom(k, n, p, 1 - p)
  if (k === n || p === 0) return 1
  return p === 1 ? 0 : ibeta(1 - p, p, n - k, k + 1)[0]
}

export function negbinomDist(f: number, s: number, p: number, cumulative: boolean): V {
  if (f < 0 || s < 1 || p <= 0 || p >= 1) return E.NUM
  return cumulative ? ibeta(p, 1 - p, s, f + 1)[0] : (s / (s + f)) * dbinom(s, s + f, p, 1 - p)
}

/** P(X = k) for k successes in a sample of n from N items with K successes (R's dhyper by binomial densities). */
function hypgeomPdf(k: number, n: number, K: number, N: number): number {
  if (k < Math.max(0, n - N + K) || k > Math.min(n, K)) return 0
  const p = n / N
  const q = (N - n) / N
  return (dbinom(k, K, p, q) * dbinom(n - k, N - K, p, q)) / dbinom(n, N, p, q)
}

export function hypgeomDist(k: number, n: number, K: number, N: number, cumulative: boolean): V {
  if (k < 0 || n < 0 || K < 0 || N <= 0 || n > N || K > N) return E.NUM
  if (!cumulative) return hypgeomPdf(k, n, K, N)
  const lo = Math.max(0, n - N + K)
  const hi = Math.min(n, K)
  if (k < lo) return 0
  if (k >= hi) return 1
  // sum the smaller side from k outwards by the ratio of neighbouring terms, until the terms vanish
  if (k < (n * K) / N) {
    let term = hypgeomPdf(k, n, K, N)
    let acc = term
    for (let i = k; i > lo && term > acc * 1e-17; i--) {
      term *= (i * (N - K - n + i)) / ((K - i + 1) * (n - i + 1))
      acc += term
    }
    return acc
  }
  let term = hypgeomPdf(k + 1, n, K, N)
  let acc = term
  for (let i = k + 1; i < hi && term > acc * 1e-17; i++) {
    term *= ((K - i) * (n - i)) / ((i + 1) * (N - K - n + i + 1))
    acc += term
  }
  return 1 - acc
}

export function poissonDist(x: number, m: number, cumulative: boolean): V {
  if (x < 0 || m < 0) return E.NUM
  if (!cumulative) return dpois(x, m)
  return m === 0 ? 1 : gammaPQ(x + 1, m)[1]
}

export function permut(n: number, k: number): V {
  if (n < 0 || k < 0 || n < k || n >= MAX_COUNT) return E.NUM
  let r = 1
  for (let i = 0; i < k && Number.isFinite(r); i++) r *= n - i
  return r
}

export const permutationA = (n: number, k: number): V => (n < 0 || k < 0 || n >= MAX_COUNT ? E.NUM : n ** k)

export const functions: Record<string, FnDef> = {
  'BETA.DIST': { args: 'n n n b n? n?', fn: betaDist },
  'BINOM.DIST': { args: 'i i n b', fn: binomDist },
  'BINOM.DIST.RANGE': {
    args: 'i n i i?',
    fn: (n: number, p: number, s: number, s2 = s) => {
      if (n < 0 || p < 0 || p > 1 || s < 0 || s > n || s2 < s || s2 > n) return E.NUM
      let r = 0
      for (let k = s; k <= s2; k++) r += combin(n, k) * Math.pow(p, k) * Math.pow(1 - p, n - k)
      return r
    },
  },
  'CHISQ.DIST': { args: 'n i b', fn: chisqDist },
  'CHISQ.DIST.RT': { args: 'n i', fn: chisqDistRt },
  'EXPON.DIST': { args: 'n n b', fn: exponDist },
  'F.DIST': { args: 'n i i b', fn: fDist },
  'F.DIST.RT': { args: 'n i i', fn: fDistRt },
  GAMMA: { args: 'n', fn: (x: number) => (x <= 0 && Number.isInteger(x) ? E.NUM : gammafn(x)) },
  'GAMMA.DIST': { args: 'n n n b', fn: gammaDist },
  GAMMALN: { args: 'n', fn: (x: number) => (x <= 0 ? E.NUM : gammaln(x)) },
  'GAMMALN.PRECISE': { args: 'n', fn: (x: number) => (x <= 0 ? E.NUM : gammaln(x)) },
  GAUSS: { args: 'n', fn: (z: number) => normCdf(z)[0] - 0.5 },
  'HYPGEOM.DIST': { args: 'i i i i b', fn: hypgeomDist },
  'LOGNORM.DIST': { args: 'n n n b', fn: lognormDist },
  'NEGBINOM.DIST': { args: 'i i n b', fn: negbinomDist },
  'NORM.DIST': { args: 'n n n b', fn: normDist },
  'NORM.S.DIST': { args: 'n b', fn: normSDist },
  PERMUT: { args: 'i i', fn: permut },
  PERMUTATIONA: { args: 'i i', fn: permutationA },
  PHI: { args: 'n', fn: normPdf },
  'POISSON.DIST': { args: 'i n b', fn: poissonDist },
  PROB: {
    args: 'M M n n?',
    fn: (xs: Grid, ps: Grid, lo: number, hi = lo) => {
      if (xs.rows * xs.cols !== ps.rows * ps.cols) return E.NA
      let total = 0
      let r = 0
      for (let i = 0, n = Math.max(xs.h * xs.cols, ps.h * ps.cols); i < n; i++) {
        const x = gridAt(xs, i)
        const p = gridAt(ps, i)
        if (x === null && p === null) continue
        if (typeof x !== 'number' || typeof p !== 'number') return E.VALUE
        if (p <= 0 || p > 1) return E.NUM
        total += p
        if (x >= lo && x <= hi) r += p
      }
      return Number(total.toPrecision(15)) === 1 ? r : E.NUM
    },
  },
  'T.DIST': { args: 'n i b', fn: tDist },
  'T.DIST.2T': { args: 'n i', fn: tDist2T },
  'T.DIST.RT': { args: 'n i', fn: tDistRt },
  'WEIBULL.DIST': { args: 'n n n b', fn: weibullDist },
}
