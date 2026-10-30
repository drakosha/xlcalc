/**
 * The series of FORECAST.ETS and its kin as Excel reads it off its timeline: points of a time that is not a number
 * dropped, sorted, put into the slots of the timeline's step, points of one slot aggregated, missing slots completed.
 * Dates on one day of every k-th month make a calendar timeline, whose targets are counted in months.
 */
import { gridAt, type Grid } from '../../args.js'
import { addMonths, daysInMonth, serialOf, ymd } from '../../dates.js'
import { E, isErr, type CErr } from '../../values.js'
import { max, min, sum } from '../common.js'
import { median } from './descriptive.js'

/** A calendar timeline: dates on one day of every `months`-th month from the first. */
interface Calendar {
  months: number
  y: number
  m: number
  d: number
}

/** The series on its timeline. */
export interface Series {
  y: number[]
  first: number
  last: number
  /** The step as FORECAST.ETS.STAT reports it. */
  step: number
  calendar?: Calendar
}

/** Aggregates of the values at one slot, by Excel's numbers 1 to 7. */
function aggregate(k: number, ys: number[]): number {
  switch (k) {
    case 2:
    case 3:
      return ys.length
    case 4:
      return max(ys)
    case 5:
      return median(ys)
    case 6:
      return min(ys)
    case 7:
      return sum(ys)
    default:
      return sum(ys) / ys.length
  }
}

/** The most frequent difference between successive times, differences within 1E-8 of each other counted as one. */
function stepOf(ts: readonly number[]): number {
  const groups: [number, number][] = []
  for (let i = 1; i < ts.length; i++) {
    const d = ts[i] - ts[i - 1]
    const g = groups.find(([v]) => Math.abs(d - v) <= 1e-8 * Math.max(1, Math.abs(d)))
    if (g) g[1]++
    else groups.push([d, 1])
  }
  let best = groups[0]
  for (const g of groups) if (g[1] > best[1]) best = g
  return best[0]
}

/** The calendar timeline the dates make, if each is on one day of the month and they are a constant number of months apart. */
function calendarOf(ts: readonly number[]): Calendar | undefined {
  const dates = ts.map((t) => (t === Math.trunc(t) && t > 60 ? ymd(t) : undefined))
  if (dates.some((d) => !d || d.d !== dates[0]!.d)) return undefined
  const index = dates.map((d) => d!.y * 12 + d!.m)
  const months = index[1] - index[0]
  if (!(months > 0) || index.some((v, i) => i > 0 && v - index[i - 1] !== months)) return undefined
  return { months, ...dates[0]! }
}

/** The serial of the calendar timeline's date `k` steps from its first. */
function calendarAt(c: Calendar, k: number): number {
  const { y, m } = addMonths(c.y, c.m, k * c.months)
  return serialOf(y, m, Math.min(c.d, daysInMonth(y, m)))
}

/**
 * The series of the values on the timeline: a point with a time that is not a number dropped, a value that is not one
 * missing, sorted, points of one slot of the step aggregated, missing slots completed by the line between their
 * neighbours (or 0), #NUM! past 30 % of the slots.
 */
export function seriesOf(values: Grid, timeline: Grid, complete: boolean, aggregation: number): Series | CErr {
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
  points.sort((a, b) => a[0] - b[0])
  // completed, points without a value at either end are no part of the series; without, they are zeros
  let lo = 0
  let hi = points.length
  while (complete && lo < hi && points[lo][1] === null) lo++
  while (complete && hi > lo && points[hi - 1][1] === null) hi--
  const kept = points.slice(lo, hi)
  if (kept.length < 2) return E.DIV0
  const ts = [...new Set(kept.map((p) => p[0]))]
  if (ts.length < 2) return E.NUM
  const first = ts[0]
  const last = ts[ts.length - 1]
  const calendar = calendarOf(ts)
  const step = calendar ? ts[1] - ts[0] : stepOf(ts)
  // slots counted back from the last point, a point halfway between two in the later one
  const back = calendar ? new Map(ts.map((t, i) => [t, ts.length - 1 - i])) : undefined
  const slots = new Map<number, number[]>()
  let top = 0
  for (const [t, y] of kept) {
    const k = back ? back.get(t)! : Math.floor((last - t) / step + 0.5 - 1e-8)
    top = Math.max(top, k)
    if (!slots.has(k)) slots.set(k, [])
    if (y !== null) slots.get(k)!.push(y)
  }
  const y: (number | null)[] = []
  for (let k = top; k >= 0; k--) {
    const at = slots.get(k)
    y.push(at?.length ? aggregate(aggregation, at) : null)
  }
  if (y.filter((v) => v === null).length > 0.3 * y.length) return E.NUM
  return { y: completed(y, complete), first, last, step, calendar }
}

/** Missing values by the line between the nearest values on each side (the nearest one past the ends), or 0. */
function completed(y: readonly (number | null)[], complete: boolean): number[] {
  return y.map((v, i) => {
    if (v !== null) return v
    if (!complete) return 0
    let a = i - 1
    while (a >= 0 && y[a] === null) a--
    let b = i + 1
    while (b < y.length && y[b] === null) b++
    if (a < 0) return y[b]!
    if (b >= y.length) return y[a]!
    return y[a]! + ((y[b]! - y[a]!) * (i - a)) / (b - a)
  })
}

/**
 * The steps from the last point to a target: within 0.001 of a whole number, that number; on a calendar timeline the
 * whole months and the fraction of the month the target falls in. #NUM! before the first point.
 */
export function horizonOf(s: Series, target: number): number | CErr {
  if (target < s.first) return E.NUM
  const c = s.calendar
  if (c) {
    const last = s.y.length - 1
    let q = Math.floor((target - s.last) / (30 * c.months))
    while (calendarAt(c, last + q) > target) q--
    while (calendarAt(c, last + q + 1) <= target) q++
    const a = calendarAt(c, last + q)
    return target === a ? q : q + (target - a) / (calendarAt(c, last + q + 1) - a)
  }
  const h = (target - s.last) / s.step
  const r = Math.round(h)
  return Math.abs(h - r) * 1000 <= 1 ? r : h
}
