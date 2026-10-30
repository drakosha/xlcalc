/**
 * FORECAST.ETS, FORECAST.ETS.CONFINT, FORECAST.ETS.SEASONALITY and FORECAST.ETS.STAT on series that tell the parts of
 * the model apart, measured to fit it: series without a season (seasonality 0), whose forecasts at two horizons give
 * the final level and trend and so, with Excel's own alpha and beta, its start; series with a season given, whose
 * forecasts over one season and one step more give the final seasonal states; series to detect a season in; and how
 * the timeline and values are read (order, points at one time, gaps, blanks, steps, targets, every argument's range).
 * Answers: Excel for Windows's in `test/excel/win/cases/ets-fit.json`. Excel for Mac answered #NUM! at random in two runs
 * (`docs/development.md`, "FORECAST.ETS"), which stay in `bench/coverage/ets/mac/` as a measurement. `PREDICTED`, the
 * first model's answers, was committed before Excel ran it.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'

const cells: Record<string, number | string | boolean | { formula: string }> = {}
let col = 0

const letters = (i: number): string => (i < 26 ? String.fromCharCode(65 + i) : letters(Math.floor(i / 26) - 1) + letters(i % 26))

/** A series on sheet E: its timeline and values in two new columns from row 1; the two ranges. */
function series(t: readonly (number | string | boolean | null | { formula: string })[], y: readonly (number | string | boolean | null | { formula: string })[]): [string, string] {
  const [a, b] = [letters(col), letters(col + 1)]
  col += 2
  const n = Math.max(t.length, y.length)
  for (let i = 0; i < n; i++) {
    if (t[i] !== null && t[i] !== undefined) cells[`${a}${i + 1}`] = t[i]!
    if (y[i] !== null && y[i] !== undefined) cells[`${b}${i + 1}`] = y[i]!
  }
  return [`E!${a}1:${a}${t.length}`, `E!${b}1:${b}${y.length}`]
}

const steps = (n: number, from = 1, step = 1) => Array.from({ length: n }, (_, i) => from + i * step)

/** Small noise, the same for every series: exact in binary to a few digits. */
const NOISE = [0.3, -0.2, 0.1, 0, -0.4, 0.2, 0.3, -0.1, 0.2, -0.3, 0, 0.1, -0.2, 0.4, -0.1, 0, 0.25, -0.15, 0.05, -0.35]
const noise = (i: number) => NOISE[i % NOISE.length] * (1 + Math.floor(i / NOISE.length) / 2)
const round = (v: number) => Math.round(v * 1e6) / 1e6

/** A seasonal series: base + slope·t + the pattern + noise·scale at t = 1…n. */
const seasonal = (n: number, base: number, slope: number, pattern: readonly number[], scale = 1) =>
  steps(n).map((t, i) => round(base + slope * t + pattern[i % pattern.length] + noise(i) * scale))

const HOLT: Record<string, number[]> = {
  trend: [3, 5, 9, 11, 15, 17, 21, 24, 26, 30],
  level: [12, 15, 11, 14, 18, 13, 16, 19, 15, 17, 20, 18],
  line: steps(8).map((t) => 2 * t + 1),
  flat: [5, 5, 5, 5, 5, 5],
  walk: [10, 11.5, 10.8, 12.1, 13, 12.4, 13.9, 15.2, 14.6, 15.8, 17.1, 16.5, 18, 19.2, 18.7, 20.1, 21.4, 20.9, 22.3, 23.5],
  down: [50, 48, 47, 44, 43, 41, 38, 38, 35, 33, 32, 29, 28, 26, 23],
  three: [1, 3, 2],
  four: [4, 1, 3, 6],
}

const SEASONAL: Record<string, [m: number, y: number[]]> = {
  q16: [4, seasonal(16, 20, 0.5, [3, -1, -4, 2])],
  q24: [4, seasonal(24, 100, -1.2, [10, 4, -6, -8], 2)],
  m36: [12, seasonal(36, 50, 0.3, [5, 7, 9, 6, 2, -1, -4, -6, -8, -5, -3, -2])],
  q8: [4, seasonal(8, 10, 1, [2, -2, 1, -1])],
  q10: [4, seasonal(10, 10, 1, [2, -2, 1, -1])],
  t9: [3, seasonal(9, 30, 0, [6, -1, -5])],
  w28: [7, seasonal(28, 40, 0.2, [3, 5, 4, 2, 0, -6, -8])],
  s5: [5, seasonal(9, 10, 0, [1, 2, 3, -2, -4])],
}

/** Series to detect a season in: a pattern of each length, with a trend and noise; noise alone. */
const DETECT: Record<string, number[]> = {
  ...Object.fromEntries(steps(11, 2).map((p) => [`p${p}`, seasonal(Math.max(4 * p, 24), 10, 0.1, steps(p).map((k) => ((k * 7) % p) - (p - 1) / 2))])),
  p12sine: steps(48).map((t, i) => round(100 + 10 * Math.sin((2 * Math.PI * t) / 12) + noise(i))),
  p7long: seasonal(50, 20, 0, [1, 3, 2, 0, -1, -3, -2], 0.5),
  noise: steps(30).map((_, i) => round(10 + noise(i * 7) * 3)),
  p4flat: seasonal(16, 5, 0, [1, -1, 2, -2], 0),
}

/** Months as dates: the first of each month of 2020 and 2021. */
const MONTHS = steps(24, 0).map((k) => Date.UTC(2020, k, 1) / 86400000 + 25569)

function build(): Probe[] {
  const { out, add } = collector()
  const ETS = 'FORECAST.ETS', CI = 'FORECAST.ETS.CONFINT', SEA = 'FORECAST.ETS.SEASONALITY', STAT = 'FORECAST.ETS.STAT'

  // without a season: every statistic, three horizons, the interval at three horizons and three levels; then automatic
  for (const [name, y] of Object.entries(HOLT)) {
    const [t, v] = series(steps(y.length), y)
    const last = y.length
    add(`holt-${name}`, STAT, ...steps(8).map((k) => `${STAT}(${v},${t},${k},0)`))
    add(`holt-${name}`, ETS, ...[1, 2, 3].map((h) => `${ETS}(${last + h},${v},${t},0)`))
    add(`holt-${name}`, CI, ...[1, 2, 4].map((h) => `${CI}(${last + h},${v},${t},0.95,0)`), `${CI}(${last + 1},${v},${t},0.5,0)`, `${CI}(${last + 1},${v},${t},0.99,0)`)
    add(`holt-${name}`, SEA, `${SEA}(${v},${t})`)
    add(`holt-${name}`, ETS, `${ETS}(${last + 1},${v},${t})`)
    add(`holt-${name}`, STAT, ...[1, 2, 3, 7].map((k) => `${STAT}(${v},${t},${k})`))
  }

  // a season given: every statistic, one season of forecasts and one more, the interval; then automatic
  for (const [name, [m, y]] of Object.entries(SEASONAL)) {
    const [t, v] = series(steps(y.length), y)
    const last = y.length
    add(`aaa-${name}`, STAT, ...steps(8).map((k) => `${STAT}(${v},${t},${k},${m})`))
    add(`aaa-${name}`, ETS, ...steps(m + 1).map((h) => `${ETS}(${last + h},${v},${t},${m})`))
    add(`aaa-${name}`, CI, ...[1, 2, m + 1].map((h) => `${CI}(${last + h},${v},${t},0.95,${m})`))
    add(`aaa-${name}`, SEA, `${SEA}(${v},${t})`)
    add(`aaa-${name}`, ETS, `${ETS}(${last + 1},${v},${t})`)
    add(`aaa-${name}`, STAT, ...[1, 2, 3, 7].map((k) => `${STAT}(${v},${t},${k})`))
  }

  // detecting a season
  for (const [name, y] of Object.entries(DETECT)) {
    const [t, v] = series(steps(y.length), y)
    add(`detect-${name}`, SEA, `${SEA}(${v},${t})`)
    add(`detect-${name}`, ETS, `${ETS}(${y.length + 1},${v},${t})`)
  }

  const trend = HOLT.trend
  // the timeline out of order, with the values beside it
  {
    const order = [3, 0, 7, 9, 1, 5, 2, 8, 4, 6]
    const [t, v] = series(order.map((i) => i + 1), order.map((i) => trend[i]))
    add('order', ETS, `${ETS}(11,${v},${t},0)`, `${ETS}(11,${v},${t})`)
    add('order', STAT, `${STAT}(${v},${t},1,0)`, `${STAT}(${v},${t},7,0)`)
  }
  // points at one time: each aggregation, out of range
  {
    const t = [1, 2, 2, 3, 4, 5, 5, 5, 6, 7, 8, 9, 10]
    const y = [3, 5, 6, 9, 11, 15, 14, 19, 17, 21, 24, 26, 30]
    const [rt, rv] = series(t, y)
    add('dup', ETS, `${ETS}(11,${rv},${rt},0)`, ...[0, 1, 2, 3, 4, 5, 6, 7, 8, -1, 1.5].map((a) => `${ETS}(11,${rv},${rt},0,1,${a})`))
    add('dup', STAT, `${STAT}(${rv},${rt},8,0)`, `${STAT}(${rv},${rt},6,0,1,7)`)
    add('dup', SEA, `${SEA}(${rv},${rt})`)
  }
  // gaps: one point, three points, completed by neighbours or zeros; the 30 % limit by the original or the full count
  {
    const cut = (drop: number[]) => steps(13).filter((k) => !drop.includes(k))
    const ys = (ts: number[]) => ts.map((k) => round(3 * k + noise(k)))
    for (const [name, drop] of Object.entries({ one: [5], run: [5, 6, 7], three: [3, 7, 11], four: [3, 5, 7, 11], five: [2, 4, 6, 8, 10] })) {
      const ts = cut(drop)
      const [t, v] = series(ts, ys(ts))
      add(`gap-${name}`, ETS, `${ETS}(14,${v},${t},0)`, `${ETS}(14,${v},${t},0,0)`, `${ETS}(14,${v},${t},0,2)`, `${ETS}(14,${v},${t},0,-1)`, `${ETS}(14,${v},${t},0,0.5)`)
      add(`gap-${name}`, STAT, `${STAT}(${v},${t},7,0)`, `${STAT}(${v},${t},7,0,0)`, `${STAT}(${v},${t},8,0)`)
    }
  }
  // values that are not numbers: a blank, text, a boolean, an error; the same in the timeline
  {
    const kinds: Record<string, number | string | boolean | null | { formula: string }> = { blank: null, text: 'x', num: '15', bool: true, err: { formula: '1/0' }, empty: { formula: '""' } }
    for (const [name, k] of Object.entries(kinds)) {
      const [t, v] = series(steps(10), trend.map((y, i) => (i === 4 ? k : y)))
      add(`value-${name}`, ETS, `${ETS}(11,${v},${t},0)`, `${ETS}(11,${v},${t},0,0)`)
      add(`value-${name}`, STAT, `${STAT}(${v},${t},7,0)`)
      const [t2, v2] = series(steps(10).map((x, i) => (i === 4 ? k : x)), trend)
      add(`time-${name}`, ETS, `${ETS}(11,${v2},${t2},0)`)
      add(`time-${name}`, STAT, `${STAT}(${v2},${t2},8,0)`)
    }
    const [t, v] = series(steps(10), [...trend.slice(0, 9), null])
    add('value-last', ETS, `${ETS}(11,${v},${t},0)`, `${ETS}(10,${v},${t},0)`)
  }
  // steps: uneven, a week, a half, an hour, months as dates, large serials, negative times
  {
    const timelines: Record<string, number[]> = {
      uneven: [1, 2, 3.5, 4.5, 5.5, 6.5, 7.5, 8.5, 9.5, 10.5],
      double: [1, 2, 4, 6, 8, 10, 12, 14, 16, 18],
      week: steps(10, 45000, 7),
      half: steps(10, 0, 0.5),
      hour: steps(10, 45000, 1 / 24),
      tenth: steps(10, 0.1, 0.1),
      months: MONTHS.slice(0, 10),
      negative: steps(10, -20, 2),
    }
    for (const [name, ts] of Object.entries(timelines)) {
      const [t, v] = series(ts, trend)
      const next = name === 'months' ? MONTHS[10] : ts[9] + (ts[9] - ts[8])
      add(`step-${name}`, STAT, `${STAT}(${v},${t},8,0)`, `${STAT}(${v},${t},1,0)`)
      add(`step-${name}`, ETS, `${ETS}(${next},${v},${t},0)`)
    }
  }
  // targets: the last point, before it, between steps, far ahead, as text, a boolean, a cell, an array
  {
    const [t, v] = series(steps(10), trend)
    for (const x of ['10', '9', '1', '10.5', '11.5', '12.25', '110', '"11"', 'TRUE', 'E!ZZ1', '{11,12}', '-5', '0'])
      add('target', x === '{11,12}' ? 'INDEX' : ETS, x === '{11,12}' ? `INDEX(${ETS}(${x},${v},${t},0),2)` : `${ETS}(${x},${v},${t},0)`)
    for (const x of ['10', '9', '10.5', '110']) add('target', CI, `${CI}(${x},${v},${t},0.95,0)`)
    // seasonality: out of range, fractional, as text, longer than half the series, at the limit
    for (const s of ['-1', '1.5', '2.5', '4.9', '"4"', 'TRUE', '5', '6', '8760', '8761', 'E!ZZ1']) add('season', ETS, `${ETS}(11,${v},${t},${s})`)
    for (const s of ['0', '1', '2', '3', '4', '5']) add('season', STAT, `${STAT}(${v},${t},3,${s})`)
    // statistic types, confidence levels
    for (const k of ['0', '9', '1.5', '7.9', '"2"', 'TRUE', '-1']) add('type', STAT, `${STAT}(${v},${t},${k},0)`)
    for (const c of ['0', '1', '0.5', '1.5', '-0.1', '0.999', '0.0001', '"0.9"', 'TRUE', 'E!ZZ1']) add('level', CI, `${CI}(11,${v},${t},${c},0)`)
    // left out arguments
    add('omit', ETS, `${ETS}(11,${v},${t},,,)`, `${ETS}(11,${v},${t},0,,)`, `${ETS}(11,${v},${t},,1,1)`)
    add('omit', CI, `${CI}(11,${v},${t},,0)`, `${CI}(11,${v},${t})`)
    add('omit', STAT, `${STAT}(${v},${t},1,,,)`)
  }
  // sizes and shapes: different counts, a row against a column, arrays, two points, one point, one time
  {
    const [t, v] = series(steps(10), trend)
    const t9 = t.replace(/10$/, '9')
    add('size', ETS, `${ETS}(11,${v},${t9},0)`, `${ETS}(11,{3,5,9,11,15,17,21,24,26,30},{1,2,3,4,5,6,7,8,9,10},0)`, `${ETS}(11,{3,5,9,11,15,17,21,24,26,30},${t},0)`)
    add('size', ETS, `${ETS}(3,{1,2},{1,2},0)`, `${ETS}(4,{1,2,4},{1,2,3},0)`, `${ETS}(2,{5},{1},0)`, `${ETS}(3,{1,2},{1,1},0)`)
    add('size', STAT, `${STAT}({1,2},{1,2},1,0)`, `${STAT}({1,2,4},{1,2,3},1,0)`, `${STAT}({1,2,4},{1,2,3},7,0)`)
    add('size', SEA, `${SEA}({1,2,4},{1,2,3})`, `${SEA}({1,2},{1,2})`)
  }
  return out
}

const PROBES = build()

export const SHEETS: SheetCells = { E: cells }

export const buildProbes = (): Probe[] => PROBES

/** The first model's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./ets-fit.predicted.json', import.meta.url), 'utf8'))
