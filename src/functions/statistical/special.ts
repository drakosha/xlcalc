/**
 * Special functions for the distributions, written here on published methods: Loader's saddle-point
 * densities (C. Loader, "Fast and Accurate Computation of Binomial Probabilities", 2000), incomplete gamma
 * and beta by series and Lentz continued fractions (Numerical Recipes) with prefactors that keep their
 * precision for large parameters, and a bracketed Newton solver for the inverses.
 */
import { deviate, erfc, gammafn } from '../jstat.js'

const LN_SQRT_2PI = 0.9189385332046728
const SQRT_2PI = 2.5066282746310002
const EPS = 1e-16
const FPMIN = 1e-300
/** Iterations allowed to a series or continued fraction: past them the answer is NaN (#NUM!), not a hang. */
const MAX_ITER = 5_000_000

/** ln Γ(x) − ((x − ½) ln x − x + ln √(2π)), the Stirling remainder. */
function stirlerr(x: number): number {
  if (x < 10) return Math.log(gammafn(x)) - (x - 0.5) * Math.log(x) + x - LN_SQRT_2PI
  const v = 1 / (x * x)
  return (1 / 12 - v * (1 / 360 - v * (1 / 1260 - v * (1 / 1680 - v * (1 / 1188 - v * (691 / 360360 - v / 156)))))) / x
}

/** ln Γ(x) for x > 0. */
export const lgamma = (x: number): number => (x < 10 ? Math.log(gammafn(x)) : (x - 0.5) * Math.log(x) - x + LN_SQRT_2PI + stirlerr(x))

/** x ln(x/m) + m − x without cancellation when x is near m (Loader). */
function bd0(x: number, m: number): number {
  if (Math.abs(x - m) < 0.1 * (x + m)) {
    let v = (x - m) / (x + m)
    let s = (x - m) * v
    let ej = 2 * x * v
    v *= v
    for (let j = 1; j < 1000; j++) {
      ej *= v
      const s1 = s + ej / (2 * j + 1)
      if (s1 === s) return s1
      s = s1
    }
  }
  return x * Math.log(x / m) + m - x
}

/** m^x e^(−m) / Γ(x + 1) for real x ≥ 0, the Poisson density. */
export function dpois(x: number, m: number): number {
  if (m === 0) return x === 0 ? 1 : 0
  if (x === 0) return Math.exp(-m)
  if (!Number.isFinite(m)) return 0
  return Math.exp(-stirlerr(x) - bd0(x, m)) / Math.sqrt(2 * Math.PI * x)
}

/** The gamma density with shape a and scale b at x > 0, through the Poisson density. */
export const dgamma = (x: number, a: number, b = 1) => (a < 1 ? (dpois(a, x / b) * a) / x : dpois(a - 1, x / b) / b)

/** C(n, x) p^x q^(n−x) for real 0 ≤ x ≤ n, with q = 1 − p given separately. */
export function dbinom(x: number, n: number, p: number, q: number): number {
  if (p === 0) return x === 0 ? 1 : 0
  if (q === 0) return x === n ? 1 : 0
  if (x === 0) return n === 0 ? 1 : Math.exp(p < 0.1 ? -bd0(n, n * q) - n * p : n * Math.log(q))
  if (x === n) return Math.exp(q < 0.1 ? -bd0(n, n * p) - n * q : n * Math.log(p))
  const lc = stirlerr(n) - stirlerr(x) - stirlerr(n - x) - bd0(x, n * p) - bd0(n - x, n * q)
  return Math.exp(lc - 0.5 * (Math.log(2 * Math.PI) + Math.log(x) + Math.log1p(-x / n)))
}

/** ln B(a, b) without the cancellation of three large ln Γ. */
export function lbeta(a: number, b: number): number {
  const p = Math.min(a, b)
  const q = Math.max(a, b)
  if (p >= 10) {
    const corr = stirlerr(p) + stirlerr(q) - stirlerr(p + q)
    return -0.5 * Math.log(q) + LN_SQRT_2PI + corr + (p - 0.5) * Math.log(p / (p + q)) + q * Math.log1p(-p / (p + q))
  }
  if (q >= 10) return lgamma(p) + stirlerr(q) - stirlerr(p + q) + p - p * Math.log(p + q) + (q - 0.5) * Math.log1p(-p / (p + q))
  return Math.log((gammafn(p) * gammafn(q)) / gammafn(p + q))
}

/** x^a y^b / B(a, b) with y = 1 − x given precisely (one of them may be close to 1). */
export function betaPrefix(x: number, y: number, a: number, b: number): number {
  if (a >= 10 && b >= 10) {
    const n = a + b
    return Math.sqrt((a * b) / (2 * Math.PI * n)) * Math.exp(stirlerr(n) - stirlerr(a) - stirlerr(b) - bd0(a, n * x) - bd0(b, n * y))
  }
  const lx = x > 0.5 ? Math.log1p(-y) : Math.log(x)
  const ly = y > 0.5 ? Math.log1p(-x) : Math.log(y)
  return Math.exp(a * lx + b * ly - lbeta(a, b))
}

/** [P(a, x), Q(a, x)], the regularized incomplete gamma functions; NaN past the iteration ceiling. */
export function gammaPQ(a: number, x: number): [number, number] {
  if (x <= 0) return [0, 1]
  if (x === Infinity) return [1, 0]
  if (x < a + 1) {
    let sum = 1
    let term = 1
    for (let n = 1; term >= sum * EPS; n++) {
      if (n > MAX_ITER) return [NaN, NaN]
      term *= x / (a + n)
      sum += term
    }
    const p = dpois(a, x) * sum
    return [p, 1 - p]
  }
  let b = x + 1 - a
  let c = 1 / FPMIN
  let d = 1 / b
  let h = d
  for (let i = 1; ; i++) {
    if (i > MAX_ITER) return [NaN, NaN]
    const an = -i * (i - a)
    b += 2
    d = an * d + b
    if (Math.abs(d) < FPMIN) d = FPMIN
    c = b + an / c
    if (Math.abs(c) < FPMIN) c = FPMIN
    d = 1 / d
    const del = d * c
    h *= del
    if (Math.abs(del - 1) < EPS) break
  }
  const q = a * dpois(a, x) * h
  return [1 - q, q]
}

/** The continued fraction of the incomplete beta function by Lentz's method; NaN past the ceiling. */
function betacf(x: number, a: number, b: number): number {
  const qab = a + b
  const qap = a + 1
  const qam = a - 1
  let c = 1
  let d = 1 - (qab * x) / qap
  if (Math.abs(d) < FPMIN) d = FPMIN
  d = 1 / d
  let h = d
  // one term of the fraction: the factor it brings to h
  const step = (aa: number) => {
    d = 1 + aa * d
    if (Math.abs(d) < FPMIN) d = FPMIN
    c = 1 + aa / c
    if (Math.abs(c) < FPMIN) c = FPMIN
    d = 1 / d
    return d * c
  }
  for (let m = 1; ; m++) {
    if (m > MAX_ITER) return NaN
    const m2 = 2 * m
    h *= step((m * (b - m) * x) / ((qam + m2) * (a + m2)))
    const del = step((-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2)))
    h *= del
    if (Math.abs(del - 1) < EPS) return h
  }
}

/** Whether the large-a expansion I_x(a, b) ≈ Q(b, −(a + (b − 1)/2) ln x) is good to 1e-12 (error ~0.1·b³/a²). */
const largeA = (a: number, b: number) => a >= 1e5 && 0.1 * Math.max(b, 1) ** 3 <= 1e-12 * a * a

/** [I_x(a, b), 1 − I_x(a, b)] with y = 1 − x given precisely. */
export function ibeta(x: number, y: number, a: number, b: number): [number, number] {
  if (x <= 0) return [0, 1]
  if (y <= 0) return [1, 0]
  // a ≫ b puts the mass near 1, where the continued fraction's terms nearly cancel (Temme's expansion)
  if (largeA(a, b)) {
    const [p, q] = gammaPQ(b, -(a + (b - 1) / 2) * Math.log1p(-y))
    return [q, p]
  }
  if (largeA(b, a)) return gammaPQ(a, -(b + (a - 1) / 2) * Math.log1p(-x))
  const pre = betaPrefix(x, y, a, b)
  // the fraction converges fast below the mean (Numerical Recipes)
  if (x < (a + 1) / (a + b + 2)) {
    const v = (pre * betacf(x, a, b)) / a
    return [v, 1 - v]
  }
  const v = (pre * betacf(y, b, a)) / b
  return [1 - v, v]
}

export const normPdf = (z: number) => Math.exp(-0.5 * z * z) / SQRT_2PI

/** [Φ(z), 1 − Φ(z)]. */
export function normCdf(z: number): [number, number] {
  const lower = 0.5 * erfc(-z / Math.SQRT2)
  const upper = 0.5 * erfc(z / Math.SQRT2)
  return [lower, upper]
}

/**
 * Φ⁻¹(p) with q = 1 − p given precisely: jStat's erfcinv start (Abramowitz and Stegun 26.2.22), refined by
 * Halley steps on an accurate erfc until they stop changing the result.
 */
export function normInv(p: number, q = 1 - p): number {
  const pp = 2 * Math.min(p, q)
  if (pp <= 0) return NaN
  let x = 0.70711 * deviate(pp / 2)
  for (let j = 0; j < 50; j++) {
    const err = erfc(x) - pp
    const step = err / (1.1283791670955126 * Math.exp(-x * x) - x * err)
    x += step
    if (Math.abs(step) <= 1e-16 * Math.abs(x)) break
  }
  return p < q ? -Math.SQRT2 * x : Math.SQRT2 * x
}

/**
 * The root of the increasing `f` in (lo, hi) from the guess `x`: Newton steps with the derivative `df`,
 * bisection (doubling while `hi` is infinite) when a step leaves the bracket; NaN past 2,000 steps.
 */
export function solve(f: (x: number) => number, df: (x: number) => number, x: number, lo: number, hi: number): number {
  for (let i = 0; i < 2000; i++) {
    const v = f(x)
    if (Number.isNaN(v)) return NaN
    if (v === 0) return x
    if (v < 0) lo = x
    else hi = x
    let next = x - v / df(x)
    if (!(next > lo && next < hi)) next = hi === Infinity ? 2 * x : (lo + hi) / 2
    if (Math.abs(next - x) <= 1e-15 * Math.abs(next) || (hi < Infinity && hi - lo <= 1e-15 * hi)) return next
    x = next
  }
  return NaN
}
