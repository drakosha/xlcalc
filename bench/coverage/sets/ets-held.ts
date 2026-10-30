/**
 * FORECAST.ETS and its kin held out, after the model, the timeline and the detection of the season were fitted: series
 * of 3 to 200 points with seasons of 0 and 2 to 24 given and detected, daily and weekly dates with gaps and points at
 * one time under each aggregation, completed and as zeros, monthly and quarterly calendars, fractional and calendar
 * targets, intervals with and without a season, text and blanks among the values, and arguments out of range. Every
 * series comes from one seeded stream. Answers: `test/excel/cases/ets-held.json` (Excel for Mac) and
 * `test/excel/win/cases/ets-held.json` (Excel for Windows). `PREDICTED`, the engine's answers, was committed before
 * Excel ran it.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'
import { CI, day, ETS, round2, SEA, seriesSheet, STAT, steps } from './ets/sheet.js'

const { cells, series } = seriesSheet()

/** A seeded stream of [0, 1). */
function stream(seed: number): () => number {
  let x = seed
  return () => {
    x = (x * 1103515245 + 12345) % 2147483648
    return x / 2147483648
  }
}
const rnd = stream(20261030)

/** A series of n points: a level, a trend, a pattern of `period` points (none for 0) and noise, to two decimals. */
function made(n: number, period: number, noise: number): number[] {
  const pattern = steps(Math.max(period, 1)).map(() => (period ? (rnd() - 0.5) * 12 : 0))
  const level = 20 + rnd() * 30
  const trend = (rnd() - 0.4) * 0.6
  return steps(n).map((x, i) => round2(level + trend * x + pattern[i % pattern.length] + (rnd() - 0.5) * 2 * noise))
}

const SEASONS = [0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 16, 18, 20, 22, 24, 1]

function build(): Probe[] {
  const { out, add } = collector()

  // lengths 3 to 200, each with a season given or detected and a pattern of its own length or another
  const lengths = [3, 4, 5, 6, 7, 9, 11, 12, 15, 18, 20, 24, 30, 36, 45, 48, 60, 75, 90, 100, 120, 144, 160, 180, 200]
  lengths.forEach((n, k) => {
    for (const j of [0, 1]) {
      const m = SEASONS[(2 * k + j) % SEASONS.length]
      const period = j ? [4, 7, 12, 5, 3][k % 5] : m === 1 ? 6 : m
      const [t, v] = series(steps(n), made(n, period, 0.5 + (k % 4)))
      const g = `len-${n}-${j}`
      const reach = Math.max(m, 1)
      if (m === 1) add(g, SEA, `${SEA}(${v},${t})`)
      add(g, STAT, ...steps(8).map((s) => `${STAT}(${v},${t},${s},${m})`))
      add(g, ETS, ...[1, 2.5, reach + 1].map((h) => `${ETS}(${n + h},${v},${t},${m})`))
      add(g, CI, ...[1.5, reach + 0.5].map((h) => `${CI}(${n + h},${v},${t},0.9,${m})`))
    }
  })

  // daily and weekly dates with gaps and points at one time, each aggregation, completed and as zeros
  for (const [k, step] of [1, 7, 1, 7, 1, 7, 1].entries()) {
    const n = 30 + 7 * k
    const ys = made(n, k % 2 ? 0 : 7, 1)
    const ts = steps(n, 45000, step)
    const drop = new Set(steps(Math.floor(n / 10), 0).map((i) => 3 + i * 9))
    const t: number[] = []
    const y: number[] = []
    ts.forEach((x, i) => {
      if (drop.has(i)) return
      t.push(x)
      y.push(ys[i])
      if (i % 11 === 5) {
        t.push(x)
        y.push(round2(ys[i] + 3))
      }
    })
    const [rt, rv] = series(t, y)
    const agg = k + 1
    const last = ts[n - 1]
    for (const c of [1, 0]) {
      const g = `dates-${step}-${agg}-${c}`
      const opts = `,${c},${agg}`
      add(g, SEA, `${SEA}(${rv},${rt}${opts})`)
      add(g, STAT, ...[1, 2, 3, 7, 8].map((s) => `${STAT}(${rv},${rt},${s},1${opts})`))
      add(g, ETS, ...[1, 1.5, 4.25].map((h) => `${ETS}(${last + h * step},${rv},${rt},1${opts})`), `${ETS}(${last + step},${rv},${rt},0${opts})`)
      add(g, CI, `${CI}(${last + 2.5 * step},${rv},${rt},0.95,1${opts})`)
    }
  }

  // monthly and quarterly calendars, seasons given and detected, calendar and in-between targets
  for (const [k, [months, n, m]] of ([[1, 36, 12], [1, 48, 1], [1, 25, 0], [3, 20, 4], [3, 28, 1], [1, 30, 6]] as const).entries()) {
    const d = [1, 15, 28, 1, 10, 31][k]
    // the 31st is each month's last day
    const ts = steps(n, 0).map((i) => (d === 31 ? day(2018, months * i + 1, 0) : day(2018, months * i, d)))
    const [t, v] = series(ts, made(n, months === 1 ? 12 : 4, 0.8))
    const g = `cal-${months}-${n}-${m}`
    const end = new Date((ts[n - 1] - 25569) * 86400000)
    const next = (q: number, dd: number) => day(end.getUTCFullYear(), end.getUTCMonth() + q * months, dd)
    if (m === 1) add(g, SEA, `${SEA}(${v},${t})`)
    add(g, STAT, ...[1, 2, 3, 6, 8].map((s) => `${STAT}(${v},${t},${s},${m})`))
    add(g, ETS, ...[[1, end.getUTCDate()], [1, 20], [3, 5], [-2, 12]].map(([q, dd]) => `${ETS}(${next(q, dd)},${v},${t},${m})`))
    add(g, CI, `${CI}(${next(2, 9)},${v},${t},0.8,${m})`)
  }

  // text, blanks and booleans among the values, completed and as zeros
  {
    const ys: (number | string | boolean | null | { formula: string })[] = made(26, 4, 0.7)
    ys[3] = 'n/a'
    ys[8] = null
    ys[9] = null
    ys[14] = true
    ys[20] = { formula: '""' }
    const [t, v] = series(steps(26), ys)
    for (const c of [1, 0]) {
      add(`kinds-${c}`, SEA, `${SEA}(${v},${t},${c})`)
      add(`kinds-${c}`, STAT, ...[1, 3, 7].map((s) => `${STAT}(${v},${t},${s},1,${c})`))
      add(`kinds-${c}`, ETS, `${ETS}(28,${v},${t},4,${c})`, `${ETS}(27.5,${v},${t},1,${c})`)
    }
  }

  // arguments out of range and of other kinds
  {
    const [t, v] = series(steps(16), made(16, 4, 0.5))
    for (const s of ['-1', '2.5', '8784', '8785', '"3"', 'FALSE', '11', '12']) add('arg-season', STAT, `${STAT}(${v},${t},3,${s})`)
    for (const c of ['2', '-1', '0.5', 'TRUE', '"1"']) add('arg-completion', ETS, `${ETS}(17,${v},${t},1,${c})`)
    for (const a of ['0', '8', '3.5', '"7"']) add('arg-aggregation', ETS, `${ETS}(17,${v},${t},1,1,${a})`)
    for (const s of ['0', '9', '8.5', '"5"']) add('arg-statistic', STAT, `${STAT}(${v},${t},${s})`)
    for (const l of ['1', '0', '-0.1', '0.999999', '"0.5"']) add('arg-level', CI, `${CI}(17,${v},${t},${l})`)
    for (const x of ['0', '0.5', '1', '16', '15.5', '"17"', '{17,18}']) add('arg-target', ETS, x === '{17,18}' ? `INDEX(${ETS}(${x},${v},${t}),2)` : `${ETS}(${x},${v},${t})`)
    for (const x of ['16', '15', '16.0001', '16.002']) add('arg-target', CI, `${CI}(${x},${v},${t})`)
    const [t2, v2] = series([...steps(15), 'x'], made(16, 4, 0.5))
    add('arg-timeline', ETS, `${ETS}(17,${v2},${t2})`)
    add('arg-size', ETS, `${ETS}(17,${v},${t.replace(/16$/, '15')})`)
    const [t3, v3] = series(steps(16), [...made(15, 4, 0.5), { formula: 'NA()' }])
    add('arg-error', ETS, `${ETS}(17,${v3},${t3})`)
  }
  return out
}

const PROBES = build()

export const SHEETS: SheetCells = { E: cells }

export const buildProbes = (): Probe[] => PROBES

/** The engine's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./ets-held.predicted.json', import.meta.url), 'utf8'))

/** The predictions Excel did not confirm: the step of quarterly dates from January (a known difference, `etsQuarterStep`). */
export const MISSED: readonly string[] = ['cal-3-20-4#5', 'cal-3-28-1#6']
