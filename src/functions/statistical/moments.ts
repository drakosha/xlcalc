/** Compensated, corrected two-pass moments; exact binary sums for ill-conditioned or out-of-scale inputs. */
import { spend } from '../../exec/budget.js'
import { exactMoments, predict, ratio, rootRatio, type ExactMoments } from './exact-moments.js'

class Sum {
  private sum = 0
  private correction = 0
  add(x: number) {
    const next = this.sum + x
    this.correction += Math.abs(this.sum) >= Math.abs(x) ? (this.sum - next) + x : (x - next) + this.sum
    this.sum = next
  }
  get value() { return this.sum + this.correction }
}

export function stableMean(xs: readonly number[]): number {
  if (!xs.length) return NaN
  const sum = new Sum()
  let constant = true
  for (const x of xs) { spend(1); sum.add(x); constant &&= x === xs[0] }
  if (constant) return xs[0]
  if (Number.isFinite(sum.value)) return sum.value / xs.length
  const exact = exactMoments(xs)
  return exact ? ratio(exact.sx, exact.n, exact.ex) : NaN
}

// Keep a substantial margin before the mean's rounding becomes significant relative to the spread.
const unsafe = (ss: number, mean: number, n: number, distinct: boolean) =>
  !Number.isFinite(ss) || ss < 0 || (distinct && (ss < 1e-280 || Math.abs(mean) * Number.EPSILON > Math.sqrt(ss / n) * 1e-8))

function central(xs: readonly number[]): { ss: number; unsafe: boolean } {
  const mean = stableMean(xs), squares = new Sum(), deviations = new Sum()
  let distinct = false
  for (const x of xs) {
    spend(1)
    const d = x - mean
    deviations.add(d); squares.add(d * d); distinct ||= x !== xs[0]
  }
  const ss = squares.value - (deviations.value / xs.length) * deviations.value
  return { ss, unsafe: unsafe(ss, mean, xs.length, distinct) }
}

/** divisor=1 gives DEVSQ; n or n-1 gives variance, optionally its root. */
export function momentValue(xs: readonly number[], divisor: number, root = false): number {
  if (!xs.length) return 0
  const m = central(xs)
  if (!m.unsafe) return root ? Math.sqrt(m.ss / divisor) : m.ss / divisor
  const exact = exactMoments(xs)
  if (!exact) return NaN
  const denominator = exact.n * BigInt(divisor)
  const value = ratio(exact.xx, denominator, 2 * exact.ex)
  // Excel's stored STDEV(±1E300) answers #NUM!: an overflowing variance is still an error.
  return root && Number.isFinite(value) ? rootRatio(exact.xx, denominator, 2 * exact.ex) : value
}

export interface PairMoments {
  n: number
  mx: number
  my: number
  xx: number
  yy: number
  xy: number
  unsafe: boolean
  exact(): ExactMoments | undefined
}

export function pairMoments(xs: readonly number[], ys: readonly number[]): PairMoments {
  const n = xs.length, mx = stableMean(xs), my = stableMean(ys)
  const xx = new Sum(), yy = new Sum(), xy = new Sum(), dxs = new Sum(), dys = new Sum()
  let distinctX = false, distinctY = false
  for (let i = 0; i < n; i++) {
    spend(1)
    const dx = xs[i] - mx, dy = ys[i] - my
    dxs.add(dx); dys.add(dy); xx.add(dx * dx); yy.add(dy * dy); xy.add(dx * dy)
    distinctX ||= xs[i] !== xs[0]; distinctY ||= ys[i] !== ys[0]
  }
  const sxx = xx.value - (dxs.value / n) * dxs.value
  const syy = yy.value - (dys.value / n) * dys.value
  const sxy = xy.value - (dxs.value / n) * dys.value
  let exact: ExactMoments | undefined
  return {
    n, mx, my, xx: sxx, yy: syy, xy: sxy,
    unsafe: unsafe(sxx, mx, n, distinctX) || unsafe(syy, my, n, distinctY) || !Number.isFinite(sxy) || (sxx > 0 && syy > 0 && Math.abs(sxy) < Math.sqrt(sxx) * Math.sqrt(syy) * 1e-8),
    exact: () => exact ??= exactMoments(xs, ys),
  }
}

export function covariance(f: PairMoments, divisor: number): number {
  if (!f.unsafe) return f.xy / divisor
  const e = f.exact()
  return e ? ratio(e.xy, e.n * BigInt(divisor), e.ex + e.ey) : NaN
}

/** Null means no variation in x (or y for correlation). NaN remains a numeric error. */
export function slopeOf(f: PairMoments): number | null {
  if (!f.unsafe) return f.xx === 0 ? null : f.xy / f.xx
  const e = f.exact()
  return !e ? NaN : e.xx === 0n ? null : ratio(e.xy, e.xx, e.ey - e.ex)
}

export function correlation(f: PairMoments): number | null {
  if (!f.unsafe) return f.xx === 0 || f.yy === 0 ? null : Math.max(-1, Math.min(1, f.xy / Math.sqrt(f.xx) / Math.sqrt(f.yy)))
  const e = f.exact()
  return !e ? NaN : e.xx === 0n || e.yy === 0n ? null : (e.xy < 0n ? -1 : 1) * rootRatio(e.xy * e.xy, e.xx * e.yy)
}

export function prediction(f: PairMoments, x: number): number | null {
  if (!f.unsafe) {
    if (f.xx === 0) return null
    const delta = (f.xy / f.xx) * (x - f.mx)
    const value = f.my + delta
    if (Number.isFinite(value) && Math.abs(value) > Math.max(Math.abs(f.my), Math.abs(delta)) * 1e-3) return value
  }
  const e = f.exact()
  return !e ? NaN : e.xx === 0n ? null : predict(e, x)
}

export function residualDeviation(f: PairMoments): number | null {
  const residual = f.yy - f.xy * (f.xy / f.xx)
  if (!f.unsafe && f.xx !== 0 && Number.isFinite(residual) && residual > Math.abs(f.yy) * 1e-3) return Math.sqrt(residual / (f.n - 2))
  const e = f.exact()
  return !e ? NaN : e.xx === 0n ? null : rootRatio(e.yy * e.xx - e.xy * e.xy, e.n * e.xx * (e.n - 2n), 2 * e.ey)
}
