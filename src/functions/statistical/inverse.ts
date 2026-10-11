/**
 * Inverse distribution functions: jStat's starting points (`../jstat.ts`) refined by a bracketed Newton
 * solver on the accurate distribution functions of `./special.ts`, with a ceiling on its steps; the domains
 * are Excel's.
 */
import { gammapinvStart, ibetainvStart } from '../jstat.js'
import type { FnDef } from '../../args.js'
import { E, type V } from '../../values.js'
import { badDf, fCdf, fPdf, MAX_COUNT, tCdf, tPdf } from './distributions.js'
import { betaPrefix, dgamma, gammaPQ, ibeta, normInv, solve } from './special.js'

/** x with P(a, x) = p, q = 1 − p given precisely; 0 where the quantile underflows. */
export function gammaInv(p: number, q: number, a: number): number {
  const x0 = gammapinvStart(p, q, a)
  if (!(x0 > 0)) return 0
  const lower = p <= q
  const f = (x: number) => {
    const [P, Q] = gammaPQ(a, x)
    return lower ? P - p : q - Q
  }
  return solve(f, (x) => dgamma(x, a), x0, 0, Infinity)
}

/** z with I_z(a, b) = p, q = 1 − p. */
function betaInv(p: number, q: number, a: number, b: number): number {
  const z0 = ibetainvStart(p, q, a, b)
  const lower = p <= q
  const f = (z: number) => {
    const [I, J] = ibeta(z, 1 - z, a, b)
    return lower ? I - p : q - J
  }
  const d = (z: number) => betaPrefix(z, 1 - z, a, b) / (z * (1 - z))
  return solve(f, d, Math.min(Math.max(z0, 1e-300), 1 - 1e-16), 0, 1)
}

/** x with P(F ≤ x) = p, q = 1 − p, from the start of the beta quantile. */
function fInv(p: number, q: number, d1: number, d2: number): number {
  const z0 = Math.min(Math.max(ibetainvStart(p, q, d1 / 2, d2 / 2), 1e-300), 1 - 1e-16)
  const lower = p <= q
  const f = (x: number) => {
    const [P, Q] = fCdf(x, d1, d2)
    return lower ? P - p : q - Q
  }
  return solve(f, (x) => fPdf(x, d1, d2), (d2 * z0) / (d1 * (1 - z0)), 0, Infinity)
}

/** t > 0 with P(T > t) = q: closed forms for 1 and 2 degrees of freedom, else from the normal deviate (Cornish–Fisher). */
function tUpperInv(q: number, df: number): number {
  if (df === 1) return 1 / Math.tan(Math.PI * q)
  if (df === 2) return (1 - 2 * q) / Math.sqrt(2 * q * (1 - q))
  const z = -normInv(q)
  const t0 = Math.max(z + (z ** 3 + z) / (4 * df), 1e-3)
  return solve((t) => q - tCdf(t, df)[1], (t) => tPdf(t, df), t0, 0, Infinity)
}

/** Student's t quantile of the lower tail probability p. */
function tInv(p: number, df: number): number {
  if (p === 0.5) return 0
  return p < 0.5 ? -tUpperInv(p, df) : tUpperInv(1 - p, df)
}

/** T.INV.2T: p up to 2, past 1 the negative quantile. */
function tInv2T(p: number, df: number): V {
  if (p <= 0 || p >= 2 || badDf(df)) return E.NUM
  if (p === 1) return 0
  return p < 1 ? tUpperInv(p / 2, df) : -tUpperInv(1 - p / 2, df)
}

const unit = (p: number) => p > 0 && p < 1

const normSInv = (p: number): V => (unit(p) ? normInv(p) : E.NUM)
const normInvExcel = (p: number, m: number, sd: number): V => (unit(p) && sd > 0 ? m + sd * normInv(p) : E.NUM)
const lognormInv = (p: number, m: number, sd: number): V => (unit(p) && sd > 0 ? Math.exp(m + sd * normInv(p)) : E.NUM)

function gammaInvExcel(p: number, a: number, b: number): V {
  if (p < 0 || p >= 1 || a <= 0 || b <= 0) return E.NUM
  return p === 0 ? 0 : gammaInv(p, 1 - p, a) * b
}

const chisqInv = (p: number, df: number): V => (p < 0 || p >= 1 || badDf(df) ? E.NUM : p === 0 ? 0 : 2 * gammaInv(p, 1 - p, df / 2))
export function chisqInvRt(q: number, df: number): V {
  if (q <= 0 || q > 1 || badDf(df)) return E.NUM
  if (q === 1) return 0
  // one degree of freedom goes through the normal quantile of 1 − q/2, as in Excel (rounding included)
  if (df === 1) return normInv(1 - q / 2) ** 2
  return 2 * gammaInv(1 - q, q, df / 2)
}

const tInvExcel = (p: number, df: number): V => (unit(p) && !badDf(df) ? tInv(p, df) : E.NUM)

const fInvExcel = (p: number, d1: number, d2: number): V => (p < 0 || p >= 1 || badDf(d1) || badDf(d2) ? E.NUM : p === 0 ? 0 : fInv(p, 1 - p, d1, d2))
const fInvRt = (q: number, d1: number, d2: number): V => (q <= 0 || q > 1 || badDf(d1) || badDf(d2) ? E.NUM : q === 1 ? 0 : fInv(1 - q, q, d1, d2))

function betaInvExcel(p: number, a: number, b: number, lo = 0, hi = 1): V {
  if (!unit(p) || a <= 0 || b <= 0 || lo >= hi) return E.NUM
  return lo + betaInv(p, 1 - p, a, b) * (hi - lo)
}

/**
 * The smallest k whose cumulative binomial probability reaches alpha. Excel compares the first term,
 * q^n as exp(n ln q), strictly: BINOM.INV(10,0.5,0.5^10) is 1, BINOM.INV(3,0.5,0.125) is 0 (exp gives a bit more).
 */
function binomInv(n: number, p: number, alpha: number): V {
  if (n < 0 || n >= MAX_COUNT || !unit(p) || !unit(alpha)) return E.NUM
  const q = 1 - p
  if (Math.exp(n * Math.log(q)) > alpha) return 0
  const cdf = (k: number) => (k >= n ? 1 : ibeta(q, p, n - k, k + 1)[0])
  let k = Math.min(Math.max(Math.floor(n * p + Math.sqrt(n * p * q) * normInv(alpha)), 1), n)
  for (let i = 0; i < 1e6; i++) {
    if (k > 1 && cdf(k - 1) >= alpha) k--
    else if (cdf(k) < alpha) k++
    else return k
  }
  return E.NUM
}

/** CONFIDENCE.T: the half-width of the t interval. */
function confidenceT(alpha: number, sd: number, n: number): V {
  if (!unit(alpha) || sd <= 0 || n < 1) return E.NUM
  if (n === 1) return E.DIV0
  const t = tInv2T(alpha, n - 1)
  return typeof t === 'number' ? (t * sd) / Math.sqrt(n) : t
}

// through 1 − α/2 as Excel, rounding included: CONFIDENCE(1E-10,1,10) is 2.045029491302821
const confidenceNorm = (alpha: number, sd: number, n: number): V => (!unit(alpha) || sd <= 0 || n < 1 ? E.NUM : (normInv(1 - alpha / 2) * sd) / Math.sqrt(n))

export const functions: Record<string, FnDef> = {
  'BETA.INV': { args: 'n n n n? n?', fn: betaInvExcel },
  'BINOM.INV': { args: 'i n n', fn: binomInv },
  'CHISQ.INV': { args: 'n i', fn: chisqInv },
  'CHISQ.INV.RT': { args: 'n i', fn: chisqInvRt },
  'CONFIDENCE.NORM': { args: 'n n i', fn: confidenceNorm },
  'CONFIDENCE.T': { args: 'n n i', fn: confidenceT },
  'F.INV': { args: 'n i i', fn: fInvExcel },
  'F.INV.RT': { args: 'n i i', fn: fInvRt },
  'GAMMA.INV': { args: 'n n n', fn: gammaInvExcel },
  'LOGNORM.INV': { args: 'n n n', fn: lognormInv },
  'NORM.INV': { args: 'n n n', fn: normInvExcel },
  'NORM.S.INV': { args: 'n', fn: normSInv },
  'T.INV': { args: 'n i', fn: tInvExcel },
  'T.INV.2T': { args: 'n i', fn: tInv2T },
}
