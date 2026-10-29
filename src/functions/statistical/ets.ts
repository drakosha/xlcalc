/**
 * FORECAST.ETS and its kin: the additive exponential smoothing with a trend and a season (AAA) Excel describes, over a
 * timeline of a constant step. A first model to measure Excel against: Hyndman's error-correction form, the heuristic
 * start of a classical decomposition, the parameters that minimise the squared one-step errors.
 */
import { gridAt, type FnDef, type Grid } from '../../args.js'
import { E, isErr, type CErr, type V } from '../../values.js'
import { normInv } from './special.js'

/** The series on its timeline: values at a constant step from the first point, gaps completed. */
export interface Series {
  y: number[]
  t0: number
  step: number
}

const MAX_SEASON = 8760

/** Aggregates of the values at one point of the timeline, by Excel's numbers 1 to 7 (0 is AVERAGE too). */
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

/** The series of the values on the timeline: sorted, points at one time aggregated, gaps completed (#NUM! for an uneven step). */
export function seriesOf(values: Grid, timeline: Grid, completion: number, aggregation: number): Series | CErr {
  const n = values.rows * values.cols
  if (n !== timeline.rows * timeline.cols) return E.NA
  const points: [number, number | null][] = []
  for (let i = 0; i < n; i++) {
    const t = gridAt(timeline, i)
    const y = gridAt(values, i)
    if (isErr(t)) return t
    if (isErr(y)) return y
    if (typeof t !== 'number') return E.VALUE
    points.push([t, typeof y === 'number' ? y : null])
  }
  points.sort((a, b) => a[0] - b[0])
  const ts: number[] = []
  const ys: (number | null)[] = []
  for (let i = 0; i < points.length; ) {
    let j = i
    const at: number[] = []
    for (; j < points.length && points[j][0] === points[i][0]; j++) if (points[j][1] !== null) at.push(points[j][1]!)
    ts.push(points[i][0])
    ys.push(at.length ? aggregate(aggregation, at) : null)
    i = j
  }
  if (ts.length < 2) return E.NUM
  let step = Infinity
  for (let i = 1; i < ts.length; i++) step = Math.min(step, ts[i] - ts[i - 1])
  const y: (number | null)[] = []
  for (let i = 0; i < ts.length; i++) {
    if (i) {
      const k = (ts[i] - ts[i - 1]) / step
      if (Math.abs(k - Math.round(k)) > 1e-9) return E.NUM
      for (let g = 1; g < Math.round(k); g++) y.push(null)
    }
    y.push(ys[i])
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
  return { y: full, t0: ts[0], step }
}

/** The length of the season by the largest autocorrelation of the differenced series, 1 for none. */
export function seasonOf(y: readonly number[]): number {
  const d = y.slice(1).map((v, i) => v - y[i])
  const n = d.length
  const m = d.reduce((a, b) => a + b, 0) / n
  const c0 = d.reduce((a, v) => a + (v - m) ** 2, 0)
  if (c0 === 0) return 1
  let best = 1
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
  season: number[]
  fitted: number[]
}

/** The start: a classical decomposition of the first two seasons, a line through the first two points without one. */
function start(y: readonly number[], m: number): { l: number; b: number; s: number[] } {
  if (m <= 1) return { l: y[0], b: y[1] - y[0], s: [0] }
  const avg = (a: number) => y.slice(a, a + m).reduce((p, q) => p + q, 0) / m
  const a1 = avg(0)
  const a2 = y.length >= 2 * m ? avg(m) : a1
  const b = (a2 - a1) / m
  const l = a1 - b * (m - 1) / 2
  const s = y.slice(0, m).map((v, j) => v - (l + b * j))
  return { l: l - b, b, s }
}

/** Runs the recursions with given parameters from the start; the one-step forecasts and the final states. */
function run(y: readonly number[], m: number, alpha: number, beta: number, gamma: number): Fit {
  const k = Math.max(m, 1)
  const { l: l0, b: b0, s: s0 } = start(y, m)
  let l = l0
  let b = b0
  const s = [...s0]
  const fitted: number[] = []
  for (let t = 0; t < y.length; t++) {
    const j = t % k
    const f = l + b + (m > 1 ? s[j] : 0)
    fitted.push(f)
    const e = y[t] - f
    const nl = l + b + alpha * e
    b = b + beta * e
    l = nl
    if (m > 1) s[j] += gamma * e
  }
  const season = m > 1 ? y.map((_, i) => i).slice(0, k).map((i) => s[(y.length + i) % k]) : [0]
  return { alpha, beta, gamma, m, level: l, trend: b, season, fitted }
}

const sse = (y: readonly number[], f: Fit) => f.fitted.reduce((a, v, i) => a + (y[i] - v) ** 2, 0)

/** Nelder–Mead over the parameters inside [0, 1]: the fit of the least squared one-step errors. */
export function fitOf(y: readonly number[], m: number): Fit {
  const dim = m > 1 ? 3 : 2
  const clamp = (p: number[]) => p.map((v) => Math.min(1, Math.max(0, v)))
  const cost = (p: number[]) => {
    const q = clamp(p)
    return sse(y, run(y, m, q[0], q[1], dim > 2 ? q[2] : 0))
  }
  let simplex: number[][] = [Array(dim).fill(0.2)]
  for (let i = 0; i < dim; i++) simplex.push(simplex[0].map((v, j) => (j === i ? 0.5 : v)))
  let costs = simplex.map(cost)
  for (let it = 0; it < 400; it++) {
    const order = costs.map((c, i) => i).sort((a, b) => costs[a] - costs[b])
    simplex = order.map((i) => simplex[i])
    costs = order.map((i) => costs[i])
    if (costs[dim] - costs[0] <= 1e-12 * (1 + Math.abs(costs[0]))) break
    const c = simplex[0].map((_, j) => simplex.slice(0, dim).reduce((a, p) => a + p[j], 0) / dim)
    const toward = (k: number) => c.map((v, j) => v + k * (simplex[dim][j] - v))
    const r = toward(-1)
    const cr = cost(r)
    if (cr < costs[0]) {
      const x = toward(-2)
      const cx = cost(x)
      ;[simplex[dim], costs[dim]] = cx < cr ? [x, cx] : [r, cr]
    } else if (cr < costs[dim - 1]) [simplex[dim], costs[dim]] = [r, cr]
    else {
      const k = toward(0.5)
      const ck = cost(k)
      if (ck < costs[dim]) [simplex[dim], costs[dim]] = [k, ck]
      else {
        for (let i = 1; i <= dim; i++) simplex[i] = simplex[i].map((v, j) => simplex[0][j] + (v - simplex[0][j]) / 2)
        costs = simplex.map(cost)
      }
    }
  }
  const p = clamp(simplex[0])
  return run(y, m, p[0], p[1], dim > 2 ? p[2] : 0)
}

/** The forecast h steps past the last point. */
function ahead(f: Fit, h: number): number {
  const k = Math.max(f.m, 1)
  return f.level + h * f.trend + (f.m > 1 ? f.season[(Math.ceil(h) - 1) % k] : 0)
}

/** The model of an ETS call's arguments: the series, its season and fit. */
interface Model {
  s: Series
  fit: Fit
}

function modelOf(values: Grid, timeline: Grid, season: number | undefined, completion: number | undefined, aggregation: number | undefined): Model | CErr {
  const c = completion ?? 1
  const a = aggregation ?? 0
  if (c !== 0 && c !== 1) return E.NUM
  if (!(a >= 0 && a <= 7) || a !== Math.trunc(a)) return E.NUM
  const m0 = season ?? 1
  if (m0 < 0 || m0 > MAX_SEASON || m0 !== Math.trunc(m0)) return E.NUM
  const s = seriesOf(values, timeline, c, a)
  if (isErr(s)) return s
  const m = m0 === 1 ? seasonOf(s.y) : m0
  if (m > 1 && s.y.length < 2 * m) return E.NUM
  return { s, fit: fitOf(s.y, m) }
}

/** The steps from the last point to a target date, #NUM! before it. */
function stepsTo(s: Series, target: number): number | CErr {
  const h = (target - (s.t0 + (s.y.length - 1) * s.step)) / s.step
  return h < 0 ? E.NUM : h
}

function forecastEts(target: number, values: Grid, timeline: Grid, season?: number, completion?: number, aggregation?: number): V {
  const md = modelOf(values, timeline, season, completion, aggregation)
  if (isErr(md)) return md
  const h = stepsTo(md.s, target)
  return isErr(h) ? h : ahead(md.fit, h)
}

/** The half width of the prediction interval: the h-step variance of the AAA model from the one-step one. */
function confint(target: number, values: Grid, timeline: Grid, level?: number, season?: number, completion?: number, aggregation?: number): V {
  const cl = level ?? 0.95
  if (!(cl > 0 && cl < 1)) return E.NUM
  const md = modelOf(values, timeline, season, completion, aggregation)
  if (isErr(md)) return md
  const h = stepsTo(md.s, target)
  if (isErr(h)) return h
  const { fit, s } = md
  const n = s.y.length
  const mse = fit.fitted.reduce((a, v, i) => a + (s.y[i] - v) ** 2, 0) / n
  let w = 1
  for (let j = 1; j < Math.ceil(h); j++) w += (fit.alpha + fit.beta * j + (fit.m > 1 && j % fit.m === 0 ? fit.gamma : 0)) ** 2
  return normInv((1 + cl) / 2) * Math.sqrt(mse * w)
}

function seasonality(values: Grid, timeline: Grid, completion?: number, aggregation?: number): V {
  const md = modelOf(values, timeline, 1, completion, aggregation)
  return isErr(md) ? md : md.fit.m
}

function stat(values: Grid, timeline: Grid, type: number, season?: number, completion?: number, aggregation?: number): V {
  if (!(type >= 1 && type <= 8) || type !== Math.trunc(type)) return E.NUM
  const md = modelOf(values, timeline, season, completion, aggregation)
  if (isErr(md)) return md
  const { fit, s } = md
  const y = s.y
  const n = y.length
  const errs = fit.fitted.map((v, i) => y[i] - v)
  const mae = errs.reduce((a, e) => a + Math.abs(e), 0) / n
  switch (type) {
    case 1:
      return fit.alpha
    case 2:
      return fit.beta
    case 3:
      return fit.gamma
    case 4: {
      let d = 0
      for (let i = 1; i < n; i++) d += Math.abs(y[i] - y[i - 1])
      return d ? mae / (d / (n - 1)) : E.DIV0
    }
    case 5:
      return errs.reduce((a, e, i) => a + (y[i] || fit.fitted[i] ? (2 * Math.abs(e)) / (Math.abs(y[i]) + Math.abs(fit.fitted[i])) : 0), 0) / n
    case 6:
      return mae
    case 7:
      return Math.sqrt(errs.reduce((a, e) => a + e * e, 0) / n)
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
