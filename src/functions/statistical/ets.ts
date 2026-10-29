/**
 * FORECAST.ETS and its kin: the additive exponential smoothing with a trend and a season (AAA) Excel describes, over a
 * timeline of a constant step. Measured against Excel (`bench/coverage/sets/ets-fit.ts`): the error-correction form of
 * the recursions, a start at the first value with the least-squares slope, accuracy measured over the last 40 % of a
 * series longer than 10, the interval of the AAA model's variance. The choice of the parameters is a guess still.
 */
import { gridAt, type FnDef, type Grid } from '../../args.js'
import { E, isErr, type CErr, type V } from '../../values.js'
import { normInv } from './special.js'

/** The series on its timeline: values at a constant step from the first point, gaps completed. */
export interface Series {
  y: number[]
  /** The last point of the timeline as given and the step. */
  last: number
  first: number
  step: number
}

const MAX_SEASON = 8760
/** Excel's γ of a model without a season. */
const NO_GAMMA = 2.220446049250313e-16

/** Aggregates of the values at one point of the timeline, by Excel's numbers 1 to 7. */
function aggregate(k: number, ys: number[]): number {
  switch (k) {
    case 2:
    case 3:
      return ys.length
    case 4:
      return Math.max(...ys)
    case 5: {
      const s = [...ys].sort((a, b) => a - b)
      const h = s.length >> 1
      return s.length % 2 ? s[h] : (s[h - 1] + s[h]) / 2
    }
    case 6:
      return Math.min(...ys)
    case 7:
      return ys.reduce((a, b) => a + b, 0)
    default:
      return ys.reduce((a, b) => a + b, 0) / ys.length
  }
}

/** The most frequent difference between successive times, the smallest of equals. */
function stepOf(ts: readonly number[]): number {
  const count = new Map<number, number>()
  for (let i = 1; i < ts.length; i++) count.set(ts[i] - ts[i - 1], (count.get(ts[i] - ts[i - 1]) ?? 0) + 1)
  let best = 0
  let top = 0
  for (const [d, c] of count) if (c > top || (c === top && d < best)) [best, top] = [d, c]
  return best
}

/**
 * The series of the values on the timeline: a point with a time that is not a number dropped, a value that is not one
 * missing, sorted, points of one slot of the step aggregated, missing slots completed (#NUM! past 30 %).
 */
export function seriesOf(values: Grid, timeline: Grid, completion: number, aggregation: number): Series | CErr {
  const n = values.rows * values.cols
  if (n !== timeline.rows * timeline.cols) return E.NA
  const points: [number, number | null][] = []
  for (let i = 0; i < n; i++) {
    const t = gridAt(timeline, i)
    const y = gridAt(values, i)
    if (isErr(t)) return t
    if (isErr(y)) return y
    if (typeof t === 'number') points.push([t, typeof y === 'number' ? y : null])
  }
  if (points.length < 2) return E.DIV0
  points.sort((a, b) => a[0] - b[0])
  const ts = [...new Set(points.map((p) => p[0]))]
  if (ts.length < 2) return E.NUM
  const step = stepOf(ts)
  const first = ts[0]
  const slots = new Map<number, number[]>()
  let top = 0
  for (const [t, y] of points) {
    const k = Math.round((t - first) / step)
    top = Math.max(top, k)
    if (!slots.has(k)) slots.set(k, [])
    if (y !== null) slots.get(k)!.push(y)
  }
  const y: (number | null)[] = []
  for (let k = 0; k <= top; k++) {
    const at = slots.get(k)
    y.push(at?.length ? aggregate(aggregation, at) : null)
  }
  const missing = y.filter((v) => v === null).length
  if (missing > 0.3 * y.length) return E.NUM
  const full = y.map((v, i) => {
    if (v !== null) return v
    if (!completion) return 0
    let a = i - 1
    while (a >= 0 && y[a] === null) a--
    let b = i + 1
    while (b < y.length && y[b] === null) b++
    const l = a >= 0 ? y[a]! : y[b]!
    const r = b < y.length ? y[b]! : l
    return (l + r) / 2
  })
  return { y: full, first, last: ts[ts.length - 1], step }
}

/** The length of the season by the largest autocorrelation of the differenced series, 0 for none. */
export function seasonOf(y: readonly number[]): number {
  const d = y.slice(1).map((v, i) => v - y[i])
  const n = d.length
  const m = d.reduce((a, b) => a + b, 0) / n
  const c0 = d.reduce((a, v) => a + (v - m) ** 2, 0)
  if (c0 === 0) return 0
  let best = 0
  let top = 0
  for (let lag = 2; lag <= Math.floor(y.length / 2); lag++) {
    let c = 0
    for (let i = lag; i < n; i++) c += (d[i] - m) * (d[i - lag] - m)
    const r = c / c0
    if (r > top && r > 2 / Math.sqrt(n)) {
      top = r
      best = lag
    }
  }
  return best
}

/** A fitted model: parameters, the states after the last point, the one-step forecasts in sample. */
export interface Fit {
  alpha: number
  beta: number
  gamma: number
  m: number
  level: number
  trend: number
  /** The seasonal states of the last season, by their place after the last point. */
  season: number[]
  fitted: number[]
}

/** The least-squares slope of values on their places 1…n. */
function slope(y: readonly number[]): number {
  const n = y.length
  const mx = (n + 1) / 2
  const my = y.reduce((a, b) => a + b, 0) / n
  let sxy = 0
  let sxx = 0
  for (let i = 0; i < n; i++) {
    sxy += (i + 1 - mx) * (y[i] - my)
    sxx += (i + 1 - mx) ** 2
  }
  return sxx ? sxy / sxx : 0
}

/**
 * The start, so that the first forecast is the first value: the trend the least-squares slope (of the values less their
 * season with one), the season the mean of each place of the full seasons about the line.
 */
function start(y: readonly number[], m: number): { l: number; b: number; s: number[] } {
  if (m <= 1) {
    const b = slope(y)
    return { l: y[0] - b, b, s: [0] }
  }
  const b1 = slope(y)
  const k = Math.max(1, Math.floor(y.length / m))
  const s = Array.from({ length: m }, (_, j) => {
    let a = 0
    for (let i = 0; i < k; i++) a += y[i * m + j] - b1 * (i * m + j)
    return a / k
  })
  const mean = s.reduce((a, v) => a + v, 0) / m
  for (let j = 0; j < m; j++) s[j] -= mean
  const b = slope(y.map((v, i) => v - s[i % m]))
  return { l: y[0] - s[0] - b, b, s }
}

/** Runs the recursions with given parameters from the start: the one-step forecasts and the final states. */
function run(y: readonly number[], m: number, alpha: number, beta: number, gamma: number): Fit {
  const k = Math.max(m, 1)
  const { l: l0, b: b0, s: s0 } = start(y, m)
  let l = l0
  let b = b0
  const s = [...s0]
  const fitted: number[] = []
  for (let t = 0; t < y.length; t++) {
    const j = t % k
    const f = l + b + s[j]
    fitted.push(f)
    const e = y[t] - f
    l = l + b + alpha * e
    b = b + beta * e
    if (m > 1) s[j] += gamma * e
  }
  const season = Array.from({ length: k }, (_, i) => s[(y.length + i) % k])
  return { alpha, beta, gamma: m > 1 ? gamma : NO_GAMMA, m, level: l, trend: b, season, fitted }
}

const sse = (y: readonly number[], f: Fit) => f.fitted.reduce((a, v, i) => a + (y[i] - v) ** 2, 0)

/** The candidate values of each parameter. */
const GRID = Array.from({ length: 8 }, (_, k) => 0.001 + 0.125 * k)

/** The parameters of the least squared one-step errors over the grid, the first of equals. */
export function fitOf(y: readonly number[], m: number): Fit {
  let best: Fit | undefined
  let low = Infinity
  for (const a of GRID)
    for (const b of GRID)
      for (const g of m > 1 ? GRID : [0]) {
        const f = run(y, m, a, b, g)
        const c = sse(y, f)
        if (c < low) [best, low] = [f, c]
      }
  return best!
}

/** The forecast h steps past the last point. */
function ahead(f: Fit, h: number): number {
  const k = Math.max(f.m, 1)
  return f.level + h * f.trend + f.season[(((Math.ceil(h) - 1) % k) + k) % k]
}

/** The model of an ETS call's arguments: the series and its fit. */
interface Model {
  s: Series
  fit: Fit
}

const whole = (x: number) => x === Math.trunc(x)

function modelOf(values: Grid, timeline: Grid, season = 1, completion = 1, aggregation = 1): Model | CErr {
  if (completion !== 0 && completion !== 1) return E.NUM
  if (!(aggregation >= 1 && aggregation <= 7) || !whole(aggregation)) return E.NUM
  if (season < 0 || season > MAX_SEASON || !whole(season)) return E.NUM
  const s = seriesOf(values, timeline, completion, aggregation)
  if (isErr(s)) return s
  const m = season === 1 ? seasonOf(s.y) : season < s.y.length ? season : 0
  return { s, fit: fitOf(s.y, m) }
}

/** The steps from the last point to a target date; #NUM! before the first. */
function stepsTo(s: Series, target: number): number | CErr {
  return target < s.first ? E.NUM : (target - s.last) / s.step
}

function forecastEts(target: number, values: Grid, timeline: Grid, season?: number, completion?: number, aggregation?: number): V {
  const md = modelOf(values, timeline, season, completion, aggregation)
  if (isErr(md)) return md
  const h = stepsTo(md.s, target)
  return isErr(h) ? h : ahead(md.fit, h)
}

/** The number of last one-step errors the accuracy is measured over: 40 % of a series longer than 10, else all. */
const tested = (n: number) => (n > 10 ? Math.floor(0.4 * n) : n)

/** The half width of the prediction interval: the h-step variance of the AAA model from the one-step one. */
function confint(target: number, values: Grid, timeline: Grid, level = 0.95, season?: number, completion?: number, aggregation?: number): V {
  if (!(level >= 0 && level < 1)) return E.NUM
  const md = modelOf(values, timeline, season, completion, aggregation)
  if (isErr(md)) return md
  const h = stepsTo(md.s, target)
  if (isErr(h)) return h
  if (h <= 0) return E.NUM
  const { fit, s } = md
  const mse = sse(s.y, fit) / s.y.length
  let w = Math.min(h, 1)
  for (let j = 1; j < Math.ceil(h); j++) w += (fit.alpha + fit.beta * j + (fit.m > 1 && j % fit.m === 0 ? fit.gamma : 0)) ** 2
  return normInv((1 + level) / 2) * Math.sqrt(mse * w)
}

function seasonality(values: Grid, timeline: Grid, completion?: number, aggregation?: number): V {
  const md = modelOf(values, timeline, 1, completion, aggregation)
  return isErr(md) ? md : md.fit.m
}

function stat(values: Grid, timeline: Grid, type: number, season?: number, completion?: number, aggregation?: number): V {
  if (!(type >= 1 && type <= 8) || !whole(type)) return E.NUM
  const md = modelOf(values, timeline, season, completion, aggregation)
  if (isErr(md)) return md
  const { fit, s } = md
  const y = s.y
  const n = y.length
  const k = tested(n)
  const from = n - k
  const errs = fit.fitted.slice(from).map((v, i) => y[from + i] - v)
  const mae = errs.reduce((a, e) => a + Math.abs(e), 0) / k
  switch (type) {
    case 1:
      return fit.alpha
    case 2:
      return fit.beta
    case 3:
      return fit.gamma
    case 4: {
      const train = n > 10 ? from : n
      let d = 0
      for (let i = 1; i < train; i++) d += Math.abs(y[i] - y[i - 1])
      return d ? mae / (d / (train - 1)) : 0
    }
    case 5:
      return errs.reduce((a, e, i) => a + (e ? (2 * Math.abs(e)) / (Math.abs(y[from + i]) + Math.abs(fit.fitted[from + i])) : 0), 0) / k
    case 6:
      return mae
    case 7:
      return Math.sqrt(errs.reduce((a, e) => a + e * e, 0) / k)
    default:
      return s.step
  }
}

export const functions: Record<string, FnDef> = {
  'FORECAST.ETS': { args: 'n C C n? n? n?', fn: forecastEts },
  'FORECAST.ETS.CONFINT': { args: 'n C C n? n? n? n?', fn: confint },
  'FORECAST.ETS.SEASONALITY': { args: 'C C n? n?', fn: seasonality },
  'FORECAST.ETS.STAT': { args: 'C C n n? n? n?', fn: stat },
}
