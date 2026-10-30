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

/** A calendar timeline: dates on one day of every `months`-th month, its last date's month and day. */
interface Calendar {
  months: number
  y: number
  m: number
  d: number
}

/** The series on its timeline. */
export interface Series {
  y: number[]
  /** The last point's time. */
  last: number
  /** The step as FORECAST.ETS.STAT reports it. */
  step: number
  calendar?: Calendar
}

/** Aggregates of the values at one slot, by Excel's numbers 1 to 7; `others` counts the values that are not numbers. */
function aggregate(k: number, ys: number[], others: number): number {
  switch (k) {
    case 2:
      return ys.length
    case 3:
      return ys.length + others
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

/**
 * The step of a timeline: the most frequent difference between successive times, those within `near` of the first
 * of a group counted in it, and in that group its most frequent value; undefined where two groups tie.
 */
function stepOf(ds: readonly number[], near: (d: number, first: number) => boolean): number | undefined {
  const groups: number[][] = []
  for (const d of ds) {
    const g = groups.find((g) => near(d, g[0]))
    if (g) g.push(d)
    else groups.push([d])
  }
  let best = groups[0]
  let tie = false
  for (const g of groups)
    if (g.length > best.length) [best, tie] = [g, false]
    else if (g !== best && g.length === best.length) tie = true
  if (tie) return undefined
  const count = new Map<number, number>()
  for (const d of best) count.set(d, (count.get(d) ?? 0) + 1)
  let step = best[0]
  for (const [d, c] of count) if (c > count.get(step)!) step = d
  return step
}

/** Differences of one group: equal to 1E-8 of them, or on a calendar timeline within a day. */
const sameStep = (d: number, first: number) => Math.abs(d - first) <= 1e-8 * Math.max(1, Math.abs(d))
const sameDays = (d: number, first: number) => Math.abs(d - first) <= 1

/**
 * The calendar timeline the dates make, if each is on one day of the month (or the last of a shorter month) and each is
 * a month or a quarter after the one before: its months and each date's place counted back from the last.
 */
function calendarOf(ts: readonly number[]): { c: Calendar; back: Map<number, number> } | undefined {
  const dates = ts.map((t) => (t === Math.trunc(t) && t > 60 ? ymd(t) : undefined))
  if (dates.some((d) => !d)) return undefined
  const day = Math.max(...dates.map((d) => d!.d))
  if (dates.some((d) => d!.d !== Math.min(day, daysInMonth(d!.y, d!.m)))) return undefined
  const index = dates.map((d) => d!.y * 12 + d!.m)
  const months = index[1] - index[0]
  if ((months !== 1 && months !== 3) || index.some((v, i) => i > 0 && v - index[i - 1] !== months)) return undefined
  return { c: { months, ...dates[dates.length - 1]! }, back: new Map(ts.map((t, i) => [t, ts.length - 1 - i])) }
}

/** The serial of the date `k` steps from the calendar timeline's last, its day carried past a shorter month's end as DATE does. */
function calendarAt(c: Calendar, k: number): number {
  const { y, m } = addMonths(c.y, c.m, k * c.months)
  return serialOf(y, m, 1) + c.d - 1
}

/**
 * The series of the values on the timeline: a point with a time that is not a number dropped, a value that is not one
 * missing, sorted, points of one slot of the step aggregated, missing slots completed by the line between their
 * neighbours (or 0), #NUM! past 30 % of the slots.
 */
export function seriesOf(values: Grid, timeline: Grid, complete: boolean, aggregation: number): Series | CErr {
  const n = values.rows * values.cols
  if (n !== timeline.rows * timeline.cols) return E.NA
  // a point's value: a number, null for a value of another kind, undefined for a blank
  const points: [number, number | null | undefined][] = []
  for (let i = 0; i < n; i++) {
    const t = gridAt(timeline, i)
    const y = gridAt(values, i)
    if (isErr(t)) return t
    if (isErr(y)) return y
    if (typeof t === 'number') points.push([t, typeof y === 'number' ? y : y === null ? undefined : null])
  }
  points.sort((a, b) => a[0] - b[0])
  // completed, points without a value at either end are no part of the series; without, they are zeros
  let lo = 0
  let hi = points.length
  while (complete && lo < hi && typeof points[lo][1] !== 'number') lo++
  while (complete && hi > lo && typeof points[hi - 1][1] !== 'number') hi--
  const kept = points.slice(lo, hi)
  if (kept.length < 2) return E.DIV0
  const ts = [...new Set(kept.map((p) => p[0]))]
  if (ts.length < 2) return E.NUM
  const last = ts[ts.length - 1]
  const calendar = calendarOf(ts)
  const step = stepOf(ts.slice(1).map((t, i) => t - ts[i]), calendar ? sameDays : sameStep)
  if (step === undefined) return E.NUM
  // slots counted back from the last point, a point halfway between two in the later one
  const slots = new Map<number, { ys: number[]; others: number }>()
  let top = 0
  for (const [t, y] of kept) {
    const k = calendar ? calendar.back.get(t)! : Math.floor((last - t) / step + 0.5 - 1e-8)
    top = Math.max(top, k)
    const at = slots.get(k) ?? { ys: [], others: 0 }
    slots.set(k, at)
    if (typeof y === 'number') at.ys.push(y)
    else if (y === null) at.others++
  }
  const y: (number | null)[] = []
  for (let k = top; k >= 0; k--) {
    const at = slots.get(k)
    y.push(at && (at.ys.length || (aggregation === 3 && at.others)) ? aggregate(aggregation, at.ys, at.others) : null)
  }
  if (y.filter((v) => v === null).length > 0.3 * y.length) return E.NUM
  return { y: completed(y, complete), last, step, calendar: calendar?.c }
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
 * whole months and the fraction of the month the target falls in. #NUM! before the series' first step.
 */
export function horizonOf(s: Series, target: number): number | CErr {
  const h = stepsTo(s, target)
  return s.y.length - 1 + Math.trunc(h) < 0 ? E.NUM : h
}

function stepsTo(s: Series, target: number): number {
  const c = s.calendar
  if (c) {
    let q = Math.floor((target - s.last) / (30 * c.months))
    while (calendarAt(c, q) > target) q--
    while (calendarAt(c, q + 1) <= target) q++
    const a = calendarAt(c, q)
    return target === a ? q : q + (target - a) / (calendarAt(c, q + 1) - a)
  }
  const h = (target - s.last) / s.step
  const r = Math.round(h)
  return Math.abs(h - r) * 1000 <= 1 ? r : h
}
