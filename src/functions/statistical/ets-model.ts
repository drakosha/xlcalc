/**
 * The model of FORECAST.ETS and its kin: Hyndman's additive error, trend and season (ETS(A,A,A)) with Excel's candidate
 * sets of parameters, measured. Alpha, beta and gamma are not optimized: each candidate of a fixed list is run over the
 * series and the first with the least squared errors wins. The order of the sums and products is kept as measured, so
 * the parameters and the accuracy statistics come out as Excel's to the bit.
 */
import { sumSquares } from '../common.js'

/** A model fitted to a series: the parameters as Excel reports them, the states after the last point, the accuracy. */
export interface EtsFit {
  alpha: number
  beta: number
  gamma: number
  level: number
  trend: number
  /** The seasonal states by the place of a point in the season (two for a model without one). */
  season: number[]
  /** The one-step errors over the whole series, from the first point. */
  errors: number[]
  mase: number
  smape: number
  mae: number
  rmse: number
}

/** Excel's gamma for a model without a season, and the length of its seasonal buffer there. */
const NO_GAMMA = 2.220446049250313e-16
const NO_SEASON = 2

/** The candidates without a season, (alpha, beta) in the order they are tried: the first of equals wins. */
const PLAIN = [
  0.9, 0.001, 0.75, 0.001, 0.5, 0.001, 0.25, 0.001, 0.126, 0.001, 0.1, 0.001, 0.002, 0.001, 0.833333, 0.0833333,
  0.166667, 0.0833333, 0.998, 0.099, 0.1, 0.099, 0.126, 0.125, 0.666667, 0.166667, 0.333333, 0.166667, 0.998, 0.249,
  0.25, 0.249, 0.5, 0.25, 0.998, 0.499, 0.5, 0.499, 0.666667, 0.5, 0.998, 0.749, 0.75, 0.749, 0.833333, 0.75, 0.998,
  0.899, 0.9, 0.899,
]

/** The candidates with a season, (alpha, beta, gamma) in the order they are tried. */
const SEASONAL = [
  0.751, 0.001, 0.001, 0.501, 0.001, 0.001, 0.251, 0.001, 0.001, 0.126, 0.001, 0.001, 0.998, 0.099, 0.001, 0.126,
  0.125, 0.001, 0.998, 0.249, 0.001, 0.5, 0.25, 0.001, 0.251, 0.25, 0.001, 0.998, 0.499, 0.001, 0.501, 0.5, 0.001,
  0.998, 0.749, 0.001, 0.751, 0.75, 0.001, 0.998, 0.899, 0.001, 0.833333, 0.0833333, 0.0833333, 0.166667, 0.0833333,
  0.0833333, 0.833333, 0.75, 0.0833333, 0.9, 0.001, 0.099, 0.9, 0.899, 0.099, 0.002, 0.001, 0.125, 0.666667, 0.166667,
  0.166667, 0.333333, 0.166667, 0.166667, 0.666667, 0.5, 0.166667, 0.75, 0.001, 0.249, 0.75, 0.749, 0.249, 0.251,
  0.001, 0.25, 0.002, 0.001, 0.25, 0.5, 0.25, 0.25, 0.251, 0.25, 0.25, 0.5, 0.001, 0.499, 0.5, 0.499, 0.499, 0.002,
  0.001, 0.5, 0.333333, 0.166667, 0.5, 0.25, 0.001, 0.749, 0.25, 0.249, 0.749, 0.002, 0.001, 0.75, 0.166667,
  0.0833333, 0.75, 0.1, 0.001, 0.899, 0.1, 0.099, 0.899,
]

/** The least-squares line through points (x, y): its slope and intercept, sums in the order of the points. */
export function line(xs: readonly number[], ys: readonly number[]): [number, number] {
  const n = xs.length
  if (n < 2) return [0, n ? ys[0] : 0]
  let sx = 0
  let sy = 0
  let sxx = 0
  let sxy = 0
  for (let i = 0; i < n; i++) {
    sx += xs[i]
    sy += ys[i]
    sxx += xs[i] * xs[i]
    sxy += xs[i] * ys[i]
  }
  const den = n * sxx - sx * sx
  const slope = den ? (n * sxy - sx * sy) / den : 0
  return [slope, (sy - slope * sx) / n]
}

/** The places from one number to another. */
export const places = (from: number, to: number) => Array.from({ length: to - from + 1 }, (_, i) => from + i)

/** The least-squares slope of values on their places 1…n. */
const slope = (y: readonly number[]) => line(places(1, y.length), y)[0]

/**
 * A short series extended to 3.5 seasons for the start and the choice of parameters: each new point is the first point
 * of its place in the season moved along a slope, the slope of the series without a season, the mean slope of the
 * first four points of each place with one.
 */
function extended(y: readonly number[], m: number, seasonal: boolean): number[] {
  const out = [...y]
  const target = Math.ceil(3.5 * m)
  if (out.length >= target) return out
  let d = 0
  if (!seasonal) d = slope(y)
  else {
    for (let j = 0; j < m; j++) {
      const phase: number[] = []
      for (let i = j; i < y.length && phase.length < 4; i += m) phase.push(y[i])
      d += slope(phase)
    }
    d = d / m / m
  }
  for (let i = out.length; i < target; i++) {
    const j = i % m
    out.push(y[j] + (i - j) * d)
  }
  return out
}

/** The start: the seasonal states, the trend and the level before the first point, so that its forecast is its value. */
interface Start {
  level: number
  trend: number
  season: number[]
}

/**
 * The seasonal states are the mean deviations from a moving average of two seasons over the first five seasons at
 * most, each deviation taken at a point half a season less one into its window, centred to a sum of 0. The trend is
 * the slope of the series less its season; the level makes the first forecast the first value.
 */
function startOf(y: readonly number[], m: number, seasonal: boolean): Start {
  const season = new Array<number>(m).fill(0)
  let level = y[0]
  if (seasonal) {
    const end = Math.min(y.length, 5 * m)
    const count = end - 2 * m
    const averages: number[] = []
    let window = 0
    for (let i = 0; i < end; i++) {
      if (i >= 2 * m) {
        averages.push(window / (2 * m))
        window -= y[i - 2 * m]
      }
      window += y[i]
    }
    const shift = (m >> 1) - 1
    for (let k = 0; k < m; k++) {
      let total = 0
      let cnt = 0
      for (let i = k; i < count; i += m) {
        total = total + y[i + shift] - averages[i]
        cnt++
      }
      season[(k + shift) % m] = total / cnt
    }
    level = y[0] - season[0]
    let sum = 0
    for (let j = m - 1; j >= 0; j--) sum += season[j]
    const mean = sum / m
    for (let j = 0; j < m; j++) season[j] -= mean
    level += mean
  }
  const trend = slope(y.map((v, i) => v - season[i % m]))
  return { level: level - trend, trend, season }
}

/** The points the parameters are fitted on: past them a run forecasts without learning from its errors. */
export function trainedOf(n: number, m: number): number {
  if (n <= 10) return n
  const held = Math.min(Math.trunc(0.4 * n), Math.max(n - Math.trunc(0.8 * n), 2 * m, 10))
  return n - held
}

/** A run of the recursions from a start: the final states, the one-step errors and the trend after each point. */
interface Run {
  level: number
  trend: number
  season: number[]
  errors: number[]
  fitted: number[]
  trends: number[]
}

/**
 * The error-correction recursions: the forecast l + b + s, the error e = y − f, then l ← l + b + αe, b ← b + βe and
 * s ← s + γe up to `trained` points; past them the level moves along the trend and the season repeats.
 */
function run(y: readonly number[], a: number, b: number, g: number, start: Start, trained = y.length, bound = Infinity): Run | undefined {
  const season = [...start.season]
  const m = season.length
  let level = start.level
  let trend = start.trend
  const errors: number[] = []
  const fitted: number[] = []
  const trends: number[] = []
  let loss = 0
  for (let i = 0; i < y.length; i++) {
    const j = i % m
    const f = level + trend + season[j]
    const e = y[i] - f
    loss += e * e
    if (loss > bound) return undefined
    errors.push(e)
    fitted.push(f)
    if (i < trained) {
      level = a * e + trend + level
      trend = b * e + trend
      season[j] = g * e + season[j]
    } else level += trend
    trends.push(trend)
  }
  return { level, trend, season, errors, fitted, trends }
}

/**
 * The fitted model of a series at a season of `m` points (0 for none): the candidate with the least squared errors
 * over the extended series, run again over the series itself. Where its trend wanders more than its mean (or the mean
 * is 0), the forecast continues from the last value along the starting trend and beta is reported as 0.001.
 */
export function etsFit(y: readonly number[], m: number): EtsFit {
  const seasonal = m > 0
  const k = seasonal ? m : NO_SEASON
  const long = extended(y, k, seasonal)
  const start = startOf(long, k, seasonal)
  const table = seasonal ? SEASONAL : PLAIN
  const width = seasonal ? 3 : 2
  const trained = trainedOf(long.length, k)
  let best = Infinity
  let at = 0
  for (let c = 0; c < table.length; c += width) {
    const r = run(long, table[c], table[c + 1], seasonal ? table[c + 2] : NO_GAMMA, start, trained, best)
    if (!r) continue
    const loss = sumSquares(r.errors)
    if (loss < best) {
      best = loss
      at = c
      if (loss === 0) break
    }
  }
  const alpha = table[at]
  let beta = table[at + 1]
  const gamma = seasonal ? table[at + 2] : NO_GAMMA
  const n = y.length
  const r = run(y, alpha, beta, gamma, start)!
  let { level, trend } = r
  if (wanders(r.trends, n)) {
    level = y[n - 1] - r.season[(n - 1) % k]
    trend = start.trend
    beta = 0.001
  }
  return { alpha, beta, gamma, level, trend, season: r.season, errors: r.errors, ...accuracy(y, r, trainedOf(n, k)) }
}

/** Whether the trend after the first half of a series of more than 10 points (all of a shorter one) wanders. */
function wanders(trends: readonly number[], n: number): boolean {
  const from = n > 10 ? n >> 1 : 0
  const k = n > 10 ? n >> 1 : n
  let s = 0
  let q = 0
  for (let i = from + 1; i < n; i++) {
    s += trends[i]
    q += trends[i] * trends[i]
  }
  const mean = s / k
  const dev = k > 1 ? Math.sqrt(Math.abs((q - (s * s) / (k - 1)) / (k - 1))) : 0
  return mean === 0 || Math.abs(dev / mean) > 1
}

/** MASE, SMAPE, MAE and RMSE of the one-step errors past the trained points (all of them where none are held). */
function accuracy(y: readonly number[], r: Run, trained: number) {
  const n = y.length
  const from = trained < n ? trained : 0
  const count = n - from
  let abs = 0
  let sq = 0
  let sm = 0
  for (let i = from; i < n; i++) {
    const e = r.errors[i]
    abs += Math.abs(e)
    sq += e * e
    if (y[i] || r.fitted[i]) sm += Math.abs(e) / (Math.abs(y[i]) + Math.abs(r.fitted[i]))
  }
  let d = 0
  for (let i = 1; i < trained; i++) d += Math.abs(y[i] - y[i - 1])
  const denom = trained > 1 ? d / (trained - 1) : 0
  const mae = abs / count
  return { mase: denom ? mae / denom : 0, smape: (sm * 2) / count, mae, rmse: Math.sqrt(sq / count) }
}

/** Between whole steps along the line between their values, the whole part of h taken toward 0. */
function between(at: (k: number) => number, h: number): number {
  const lo = Math.trunc(h)
  if (h === lo) return at(lo)
  const first = at(lo)
  return (h - lo) * (at(lo + 1) - first) + first
}

/** The forecast h steps past the last of n points. */
export function etsForecast(f: EtsFit, n: number, h: number): number {
  const m = f.season.length
  return between((k) => k * f.trend + f.level + f.season[(((n + k - 1) % m) + m) % m], h)
}

/**
 * The variance of the h-step forecast over that of one step, of the AAA model: 1 + Σ (α + βj + γ[j a whole number of
 * seasons])², j = 1…h − 1, 0 at no step.
 */
export function etsSpread(f: EtsFit, m: number, h: number): number {
  return between((k) => {
    let w = k > 0 ? 1 : 0
    for (let j = 1; j < k; j++) w += (f.alpha + f.beta * j + (m > 0 && j % m === 0 ? f.gamma : 0)) ** 2
    return w
  }, h)
}
