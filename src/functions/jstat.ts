/**
 * Special functions and distributions for the statistical functions.
 *
 * Ported from jStat 1.9.6 (https://github.com/jstat/jstat), src/special.js (gammaln, gammafn, factorial,
 * combination, betacf, ibeta, erf) and src/distribution.js (normal, centralF and binomial distributions,
 * betinc): Copyright (c) 2013 jStat, MIT License (full text in THIRD_PARTY.md). formula.js 4.6.1 calls
 * these through jStat; the numerical code is unchanged, only typed and taken out of the jStat object.
 */

/** Natural logarithm of the gamma function (Lanczos). */
export function gammaln(x: number): number {
  const cof = [76.18009172947146, -86.50532032941677, 24.01409824083091, -1.231739572450155, 0.1208650973866179e-2, -0.5395239384953e-5]
  let ser = 1.000000000190015
  let y = x
  let tmp = x + 5.5
  tmp -= (x + 0.5) * Math.log(tmp)
  for (let j = 0; j < 6; j++) ser += cof[j] / ++y
  return Math.log((2.5066282746310005 * ser) / x) - tmp
}

/** The gamma function. */
export function gammafn(x: number): number {
  const p = [-1.716185138865495, 24.76565080557592, -379.80425647094563, 629.3311553128184, 866.9662027904133, -31451.272968848367, -36144.413418691176, 66456.14382024054]
  const q = [-30.8402300119739, 315.35062697960416, -1015.1563674902192, -3107.771671572311, 22538.118420980151, 4755.8462775278811, -134659.9598649693, -115132.2596755535]
  let fact: number | false = false
  let n = 0
  let xden = 0
  let xnum = 0
  let y = x
  let z: number
  let res: number
  if (x > 171.6243769536076) return Infinity
  if (y <= 0) {
    res = (y % 1) + 3.6e-16
    if (res) {
      fact = ((!(y & 1) ? 1 : -1) * Math.PI) / Math.sin(Math.PI * res)
      y = 1 - y
    } else return Infinity
  }
  const yi = y
  if (y < 1) z = y++
  else z = (y -= n = (y | 0) - 1) - 1
  for (let i = 0; i < 8; ++i) {
    xnum = (xnum + p[i]) * z
    xden = xden * z + q[i]
  }
  res = xnum / xden + 1
  if (yi < y) res /= yi
  else if (yi > y) {
    for (let i = 0; i < n; ++i) {
      res *= y
      y++
    }
  }
  if (fact) res = fact / res
  return res
}

const factorial = (n: number) => (n < 0 ? NaN : gammafn(n + 1))
const factorialln = (n: number) => (n < 0 ? NaN : gammaln(n + 1))

export function combination(n: number, m: number): number {
  // past 170 the factorials overflow
  return n > 170 || m > 170 ? Math.exp(factorialln(n) - factorialln(m) - factorialln(n - m)) : factorial(n) / factorial(m) / factorial(n - m)
}

/** The continued fraction of the incomplete beta function (modified Lentz's method). */
function betacf(x: number, a: number, b: number): number {
  const fpmin = 1e-30
  const qab = a + b
  const qap = a + 1
  const qam = a - 1
  let c = 1
  let d = 1 - (qab * x) / qap
  if (Math.abs(d) < fpmin) d = fpmin
  d = 1 / d
  let h = d
  for (let m = 1; m <= 100; m++) {
    const m2 = 2 * m
    let aa = (m * (b - m) * x) / ((qam + m2) * (a + m2))
    d = 1 + aa * d
    if (Math.abs(d) < fpmin) d = fpmin
    c = 1 + aa / c
    if (Math.abs(c) < fpmin) c = fpmin
    d = 1 / d
    h *= d * c
    aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2))
    d = 1 + aa * d
    if (Math.abs(d) < fpmin) d = fpmin
    c = 1 + aa / c
    if (Math.abs(c) < fpmin) c = fpmin
    d = 1 / d
    const del = d * c
    h *= del
    if (Math.abs(del - 1.0) < 3e-7) break
  }
  return h
}

/** The regularized incomplete beta function I_x(a, b); NaN outside 0 ≤ x ≤ 1. */
export function ibeta(x: number, a: number, b: number): number {
  const bt = x === 0 || x === 1 ? 0 : Math.exp(gammaln(a + b) - gammaln(a) - gammaln(b) + a * Math.log(x) + b * Math.log(1 - x))
  if (x < 0 || x > 1) return NaN
  if (x < (a + 1) / (a + b + 2)) return (bt * betacf(x, a, b)) / a
  return 1 - (bt * betacf(1 - x, b, a)) / b
}

export function erf(x: number): number {
  const cof = [
    -1.3026537197817094, 6.4196979235649026e-1, 1.9476473204185836e-2, -9.561514786808631e-3, -9.46595344482036e-4, 3.66839497852761e-4, 4.2523324806907e-5, -2.0278578112534e-5,
    -1.624290004647e-6, 1.30365583558e-6, 1.5626441722e-8, -8.5238095915e-8, 6.529054439e-9, 5.059343495e-9, -9.91364156e-10, -2.27365122e-10, 9.6467911e-11, 2.394038e-12, -6.886027e-12,
    8.94487e-13, 3.13092e-13, -1.12708e-13, 3.81e-16, 7.106e-15, -1.523e-15, -9.4e-17, 1.21e-16, -2.8e-17,
  ]
  let isneg = false
  let d = 0
  let dd = 0
  if (x < 0) {
    x = -x
    isneg = true
  }
  const t = 2 / (2 + x)
  const ty = 4 * t - 2
  for (let j = cof.length - 1; j > 0; j--) {
    const tmp = d
    d = ty * d - dd + cof[j]
    dd = tmp
  }
  const res = t * Math.exp(-x * x + 0.5 * (cof[0] + ty * d) - dd)
  return isneg ? res - 1 : 1 - res
}

export const normalPdf = (x: number, mean: number, std: number) => Math.exp(-0.5 * Math.log(2 * Math.PI) - Math.log(std) - Math.pow(x - mean, 2) / (2 * std * std))

export const normalCdf = (x: number, mean: number, std: number) => 0.5 * (1 + erf((x - mean) / Math.sqrt(2 * std * std)))

export const centralFCdf = (x: number, df1: number, df2: number) => (x < 0 ? 0 : ibeta((df1 * x) / (df1 * x + df2), df1 / 2, df2 / 2))

export function binomialPdf(k: number, n: number, p: number): number {
  return p === 0 || p === 1 ? (n * p === k ? 1 : 0) : combination(n, k) * Math.pow(p, k) * Math.pow(1 - p, n - k)
}

/** Continued fraction of the incomplete beta function, for the binomial distribution. */
function betinc(x: number, a: number, b: number, eps: number): number {
  let a0 = 0
  let b0 = 1
  let a1 = 1
  let b1 = 1
  let m9 = 0
  let a2 = 0
  while (Math.abs((a1 - a2) / a1) > eps) {
    a2 = a1
    let c9 = (-(a + m9) * (a + b + m9) * x) / (a + 2 * m9) / (a + 2 * m9 + 1)
    a0 = a1 + c9 * a0
    b0 = b1 + c9 * b0
    m9 = m9 + 1
    c9 = (m9 * (b - m9) * x) / (a + 2 * m9 - 1) / (a + 2 * m9)
    a1 = a0 + c9 * a1
    b1 = b0 + c9 * b1
    a0 = a0 / b1
    b0 = b0 / b1
    a1 = a1 / b1
    b1 = 1
  }
  return a1 / a
}

export function binomialCdf(x: number, n: number, p: number): number {
  const eps = 1e-10
  if (x < 0) return 0
  if (x >= n) return 1
  if (p < 0 || p > 1 || n <= 0) return NaN
  x = Math.floor(x)
  const a = x + 1
  const b = n - x
  const s = a + b
  const bt = Math.exp(gammaln(s) - gammaln(b) - gammaln(a) + a * Math.log(p) + b * Math.log(1 - p))
  const betacdf = p < (a + 1) / (s + 2) ? bt * betinc(p, a, b, eps) : 1 - bt * betinc(1 - p, b, a, eps)
  return Math.round((1 - betacdf) * (1 / eps)) / (1 / eps)
}
