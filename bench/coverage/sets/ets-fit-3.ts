/**
 * The third set to fit FORECAST.ETS by, after `ets-fit` and `ets-fit-2`, every season given: the longest season Excel
 * takes for a series of n points (`FORECAST.ETS.STAT` 3, gamma, tells a season from none) and the largest seasonality it
 * accepts; the step of a timeline whose most frequent differences tie; the 30 % of missing points at its edge; calendar
 * timelines (a month missing, two points in one month, every second month, years, quarters from the fifteenth, days
 * that months lack, targets before the last date and far ahead, intervals); points at one time with text among them;
 * gaps in a seasonal series; and the interval with a season, at whole and fractional horizons, where the trend is
 * corrected after the fit. Answers of Excel for Windows: `test/excel/win/cases/ets-fit-3.json`. `PREDICTED`, the
 * engine's answers, was committed before Excel ran it.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'
import { CI, day, ETS, noise, round2, seriesSheet, STAT, steps } from './ets/sheet.js'

const { cells, series } = seriesSheet()
const Z = noise(31, 80)

/** A line with noise and a pattern of three, n points. */
const wavy = (n: number, from = 0) => steps(n).map((x, i) => round2(10 + 1.5 * x + 2 * [1, -2, 1][i % 3] + Z[(i + from) % 80]))

/** A seasonal series of n points: a line, the pattern scaled, noise. */
const seasonal = (n: number, m: number, from = 0) => {
  const pat = steps(m).map((j) => ((j * 5) % (m + 1)) - m / 2)
  return steps(n).map((x, i) => round2(25 + 0.4 * x + 2 * pat[i % m] + Z[(i + from) % 80]))
}

/** The first day of the month k months after January 2021, and another day of it. */
const month = (k: number, d = 1) => day(2021, k, d)

function build(): Probe[] {
  const { out, add } = collector()

  // the longest season taken: gamma for every season shorter than the series, then the largest seasonality accepted
  for (const n of [6, 8, 9, 10, 11, 12, 13, 14, 16, 20, 24]) {
    const [t, v] = series(steps(n), wavy(n, n))
    add(`longest-${n}`, STAT, ...steps(n - 2, 2).map((m) => `${STAT}(${v},${t},3,${m})`))
  }
  {
    const [t, v] = series(steps(40), wavy(40))
    add('longest-40', STAT, ...steps(19, 21).map((m) => `${STAT}(${v},${t},3,${m})`))
    const [t10, v10] = series(steps(10), wavy(10))
    add('largest', STAT, ...[8761, 8762, 8800, 9000, 9500, 9999].map((m) => `${STAT}(${v10},${t10},3,${m})`))
  }

  // the step where the most frequent differences tie: the first of them or the smaller
  for (const [name, ts] of Object.entries({ larger: [1, 3, 5, 6, 7], smaller: [1, 2, 3, 5, 7], three: [1, 4, 5, 6, 9], late: [2, 3, 5, 7, 8, 9, 11, 13] })) {
    const [t, v] = series(ts, ts.map((x, i) => round2(4 + 2 * x + Z[i])))
    const last = ts[ts.length - 1]
    add(`tie-${name}`, STAT, `${STAT}(${v},${t},8,0)`, `${STAT}(${v},${t},7,0)`)
    add(`tie-${name}`, ETS, `${ETS}(${last + 1},${v},${t},0)`, `${ETS}(${last + 2},${v},${t},0)`)
  }

  // missing points at the edge of 30 %: 3 of 10, 6 and 7 of 20 in runs (so the step stays 1), completed and as zeros
  for (const [name, [n, gone]] of Object.entries({ m3of10: [10, [4, 5, 6]], m6of20: [20, [5, 6, 7, 12, 13, 14]], m7of20: [20, steps(7, 8)] } as Record<string, [number, number[]]>)) {
    const ts = steps(n).filter((x) => !gone.includes(x))
    const [t, v] = series(ts, ts.map((x) => round2(5 + 1.2 * x + Z[x])))
    for (const c of [1, 0]) add(`edge-${name}`, STAT, `${STAT}(${v},${t},7,0,${c})`)
    add(`edge-${name}`, ETS, `${ETS}(${n + 1},${v},${t},0)`)
  }

  // calendar timelines
  const calendar = (group: string, ts: number[], targets: number[], opts = ',0') => {
    const [t, v] = series(ts, ts.map((_, i) => round2(30 + 0.8 * i + Z[i + 40])))
    add(group, STAT, ...[8, 1, 7].map((k) => `${STAT}(${v},${t},${k}${opts})`))
    add(group, ETS, ...targets.map((x) => `${ETS}(${x},${v},${t}${opts})`))
    return [t, v]
  }
  calendar('cal-gap', steps(24, 0).filter((k) => k !== 5).map((k) => month(k)), [month(24), month(24, 16)])
  calendar('cal-gap-zero', steps(24, 0).filter((k) => k !== 5).map((k) => month(k)), [month(24)], ',0,0')
  calendar('cal-dup', [...steps(12, 0).map((k) => month(k)), month(2)], [month(12)])
  calendar('cal-every2', steps(12, 0).map((k) => month(2 * k)), [month(24), month(24, 21)])
  calendar('cal-year', steps(12, 0).map((k) => day(2010 + k, 0, 1)), [day(2022, 0, 1), day(2022, 6, 2)])
  calendar('cal-feb15', steps(12, 0).map((k) => month(1 + 3 * k, 15)), [month(37, 15), month(38, 14)])
  calendar('cal-day30', steps(12, 0).map((k) => day(2021, k, Math.min(30, new Date(Date.UTC(2021, k + 1, 0)).getUTCDate()))), [day(2022, 0, 30)])
  calendar('cal-day31', [0, 2, 4, 6, 7, 9, 11].map((k) => day(2021, k, 31)), [day(2022, 0, 31)])
  calendar('cal-before', steps(12, 0).map((k) => month(k)), [month(10, 15), month(5), month(0), month(-1, 31), month(11)])
  calendar('cal-far', steps(12, 0).map((k) => month(k)), [month(29), month(29, 20)])
  calendar('cal-unsorted', [5, 0, 9, 3, 11, 1, 7, 2, 10, 4, 8, 6].map((k) => month(k)), [month(12)])
  calendar('cal-noon', steps(12, 0).map((k) => month(k) + 0.5), [month(12) + 0.5])
  {
    const [t, v] = calendar('cal-ci', steps(12, 0).map((k) => month(k)), [month(13)])
    add('cal-ci', CI, ...[month(12, 16), month(14), month(14, 10)].map((x) => `${CI}(${x},${v},${t},0.95,0)`))
  }
  {
    const ts = steps(36, 0).map((k) => month(k))
    const [t, v] = series(ts, seasonal(36, 12))
    add('cal-season', STAT, ...[1, 2, 3, 7].map((k) => `${STAT}(${v},${t},${k},12)`))
    add('cal-season', ETS, `${ETS}(${month(36)},${v},${t},12)`, `${ETS}(${month(41, 10)},${v},${t},12)`)
    add('cal-season', CI, `${CI}(${month(37, 16)},${v},${t},0.95,12)`)
  }

  // points at one time with text or a blank among them, by COUNT, COUNTA and AVERAGE
  {
    const t = [1, 2, 2, 3, 4, 5, 6, 6, 7, 8, 9, 10]
    const y: (number | string | null)[] = [3, 5, 'x', 9, 11, 15, 17, null, 21, 24, 26, 30]
    const [rt, rv] = series(t, y)
    for (const a of [1, 2, 3]) add(`agg-text-${a}`, ETS, `${ETS}(11,${rv},${rt},0,1,${a})`)
    add('agg-text', STAT, `${STAT}(${rv},${rt},7,0,1,3)`)
  }

  // a seasonal series with gaps, completed and as zeros
  for (const c of [1, 0]) {
    const ts = steps(16).filter((x) => x !== 6 && x !== 11)
    const y = seasonal(16, 4, 7)
    const [t, v] = series(ts, ts.map((x) => y[x - 1]))
    add(`season-gap-${c}`, STAT, ...[1, 2, 3, 7].map((k) => `${STAT}(${v},${t},${k},4,${c})`))
    add(`season-gap-${c}`, ETS, `${ETS}(17,${v},${t},4,${c})`, `${ETS}(20,${v},${t},4,${c})`)
  }

  // the interval with a season: whole and fractional horizons, levels
  for (const [m, n] of [[2, 10], [3, 12], [5, 20], [7, 21], [12, 48]]) {
    const [t, v] = series(steps(n), seasonal(n, m, m))
    const hs = [0.5, 1, 1.5, m, m + 0.5, m + 1, 2 * m, 2 * m + 1, 3 * m + 2]
    add(`ci-${m}`, STAT, ...[1, 2, 3].map((k) => `${STAT}(${v},${t},${k},${m})`))
    add(`ci-${m}`, CI, ...hs.map((h) => `${CI}(${n + h},${v},${t},0.95,${m})`), ...[0.5, 0.99].map((c) => `${CI}(${n + m + 1},${v},${t},${c},${m})`))
  }

  // the interval where the trend is corrected after the fit: random walks, seasonal walks
  for (const k of [6, 14]) {
    const z = noise(500 + k, 14)
    let s = 50
    const [t, v] = series(steps(14), z.map((d) => (s = round2(s + d * (2 + (k % 5)) + (k % 3) - 1))))
    add(`ci-walk-${k}`, STAT, ...[1, 2].map((j) => `${STAT}(${v},${t},${j},0)`))
    add(`ci-walk-${k}`, CI, ...[15, 16, 17, 19, 21.5].map((x) => `${CI}(${x},${v},${t},0.95,0)`))
    add(`ci-walk-${k}`, ETS, `${ETS}(16,${v},${t},0)`)
  }
  for (const [m, k] of [[3, 11], [4, 3], [5, 13]]) {
    const z = noise(700 + k, 30)
    const pat = steps(m).map((j) => ((j * 3) % m) - 1)
    let s = 30
    const n = 3 * m
    const [t, v] = series(steps(n), steps(n).map((_, i) => (s = round2(s + 2 * z[i])) + 2 * pat[i % m]))
    add(`ci-swalk-${m}`, STAT, ...[1, 2, 3].map((j) => `${STAT}(${v},${t},${j},${m})`))
    add(`ci-swalk-${m}`, CI, ...[1, 2, m, m + 1, 2 * m + 1].map((h) => `${CI}(${n + h},${v},${t},0.95,${m})`))
    add(`ci-swalk-${m}`, ETS, ...[1, m + 1].map((h) => `${ETS}(${n + h},${v},${t},${m})`))
  }
  return out
}

const PROBES = build()

export const SHEETS: SheetCells = { E: cells }

export const buildProbes = (): Probe[] => PROBES

/** The engine's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./ets-fit-3.predicted.json', import.meta.url), 'utf8'))

/**
 * The predictions Excel did not confirm: a season of 70 % of the points or more is none (the longest-* cases), the
 * largest seasonality is 8,784, a tie of the most frequent steps is #NUM!, a timeline every second month is no calendar,
 * a calendar's step is the most frequent of its days, a calendar's last day carries past a shorter month's end, a target
 * before the first date is taken while its step is, and COUNTA counts text.
 */
export const MISSED: readonly string[] = [
  'longest-6#3', 'longest-6#4', 'longest-8#4', 'longest-8#5', 'longest-8#6', 'longest-9#5', 'longest-9#6',
  'longest-9#7', 'longest-10#6', 'longest-10#7', 'longest-10#8', 'longest-11#6', 'longest-11#7', 'longest-11#8',
  'longest-11#9', 'longest-12#7', 'longest-12#8', 'longest-12#9', 'longest-12#10', 'longest-13#8', 'longest-13#9',
  'longest-13#10', 'longest-13#11', 'longest-14#8', 'longest-14#9', 'longest-14#10', 'longest-14#11', 'longest-14#12',
  'longest-16#10', 'longest-16#11', 'longest-16#12', 'longest-16#13', 'longest-16#14', 'longest-20#13',
  'longest-20#14', 'longest-20#15', 'longest-20#16', 'longest-20#17', 'longest-20#18', 'longest-24#15',
  'longest-24#16', 'longest-24#17', 'longest-24#18', 'longest-24#19', 'longest-24#20', 'longest-24#21',
  'longest-24#22', 'longest-40#8', 'longest-40#9', 'longest-40#10', 'longest-40#11', 'longest-40#12', 'longest-40#13',
  'longest-40#14', 'longest-40#15', 'longest-40#16', 'longest-40#17', 'longest-40#18', 'longest-40#19', 'largest#2',
  'tie-larger#1', 'tie-larger#2', 'tie-larger#3', 'tie-larger#4', 'tie-smaller#1', 'tie-smaller#2', 'tie-smaller#3',
  'tie-smaller#4', 'tie-three#1', 'tie-three#2', 'tie-three#3', 'tie-three#4', 'cal-every2#1', 'cal-every2#5',
  'cal-feb15#1', 'cal-day30#4', 'cal-before#7', 'agg-text-3#1', 'agg-text#1',
]
