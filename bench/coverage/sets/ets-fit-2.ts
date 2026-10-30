/**
 * The second set to fit FORECAST.ETS by, after `ets-fit` showed the recursions, the start of a series without a season,
 * the accuracy over the last 40 % and the interval: how Excel chooses alpha, beta and gamma (series of ten points with
 * the noise scaled, random walks, seasons of four with their pattern scaled), where the last 40 % begins (lengths 5 to
 * 16), the start of a season (one season of forecasts and one step more at lengths 8 to 20), the reading of gaps,
 * blanks and points at one time with the parameters beside them, months as a timeline, intervals between steps, the
 * longest season, and seasons to detect.
 * Answers in `test/excel/cases/ets-fit-2.json` (Excel for Mac) and `test/excel/win/cases/ets-fit-2.json` (Excel for
 * Windows). `PREDICTED`, the model's answers after `ets-fit`, was committed before Excel ran it.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'

const cells: Record<string, number | string | boolean> = {}
let col = 0

const letters = (i: number): string => (i < 26 ? String.fromCharCode(65 + i) : letters(Math.floor(i / 26) - 1) + letters(i % 26))

/** A series on sheet E: its timeline and values in two new columns from row 1 (null leaves a cell blank); the two ranges. */
function series(t: readonly (number | null)[], y: readonly (number | string | null)[]): [string, string] {
  const [a, b] = [letters(col), letters(col + 1)]
  col += 2
  for (let i = 0; i < Math.max(t.length, y.length); i++) {
    if (t[i] !== null && t[i] !== undefined) cells[`${a}${i + 1}`] = t[i]!
    if (y[i] !== null && y[i] !== undefined) cells[`${b}${i + 1}`] = y[i]!
  }
  return [`E!${a}1:${a}${t.length}`, `E!${b}1:${b}${y.length}`]
}

const steps = (n: number, from = 1, step = 1) => Array.from({ length: n }, (_, i) => from + i * step)
const round2 = (v: number) => Math.round(v * 100) / 100

/** Noise from a seed: a linear congruential stream, centred, to two decimals. */
function noise(seed: number, n: number): number[] {
  let x = seed
  return Array.from({ length: n }, () => {
    x = (x * 1103515245 + 12345) % 2147483648
    return round2((x / 2147483648 - 0.5) * 2)
  })
}

const Z = noise(7, 40)
const ETS = 'FORECAST.ETS', CI = 'FORECAST.ETS.CONFINT', SEA = 'FORECAST.ETS.SEASONALITY', STAT = 'FORECAST.ETS.STAT'

function build(): Probe[] {
  const { out, add } = collector()
  /** The parameters, the accuracy and the forecasts that fix the final states, at seasonality m. */
  const fit = (group: string, t: string, v: string, last: number, m: number, ahead: number, more: number[] = []) => {
    add(group, STAT, ...[1, 2, 3, ...more].map((k) => `${STAT}(${v},${t},${k},${m})`))
    add(group, ETS, ...steps(ahead).map((h) => `${ETS}(${last + h},${v},${t},${m})`))
  }

  // ten points without a season: a line with the noise scaled, then random walks with a drift
  for (const [k, a] of [0.02, 0.05, 0.1, 0.2, 0.35, 0.5, 0.75, 1, 1.5, 2, 3, 4, 6, 8, 12, 20].entries()) {
    const [t, v] = series(steps(10), steps(10).map((x, i) => round2(10 + 2 * x + a * Z[i])))
    fit(`line-${k}`, t, v, 10, 0, 2, [6, 7])
  }
  for (let k = 0; k < 12; k++) {
    const z = noise(100 + k, 10)
    let s = 50
    const [t, v] = series(steps(10), z.map((d) => (s = round2(s + d * (1 + k) + (k % 3) - 1))))
    fit(`walk-${k}`, t, v, 10, 0, 2, [6, 7])
  }
  // lengths 5 to 16: where the measured part begins
  for (const n of steps(12, 5)) {
    const [t, v] = series(steps(n), steps(n).map((x, i) => round2(5 + 1.5 * x + 2 * Z[i])))
    fit(`len-${n}`, t, v, n, 0, 2, [4, 5, 6, 7])
  }
  // seasons of four: the pattern and noise scaled; lengths 8 to 20 with one season of forecasts and one step more
  const PATTERN = [3, -1, -4, 2]
  for (const [k, [p, a]] of [[0.5, 1], [1, 1], [2, 1], [4, 1], [1, 0.25], [1, 3], [8, 2], [0.25, 0.5]].entries()) {
    const [t, v] = series(steps(12), steps(12).map((x, i) => round2(20 + 0.5 * x + p * PATTERN[i % 4] + a * Z[i])))
    fit(`gamma-${k}`, t, v, 12, 4, 5, [6, 7])
  }
  for (const n of [8, 9, 10, 11, 12, 13, 14, 16, 20]) {
    const [t, v] = series(steps(n), steps(n).map((x, i) => round2(20 + 0.5 * x + PATTERN[i % 4] + Z[i + 3])))
    fit(`start-${n}`, t, v, n, 4, 5, [4, 5, 6, 7])
  }
  for (const [m, n] of [[3, 9], [3, 12], [2, 8], [5, 15], [6, 18]]) {
    const pat = steps(m).map((j) => ((j * 5) % (m + 1)) - m / 2)
    const [t, v] = series(steps(n), steps(n).map((x, i) => round2(30 - 0.3 * x + 2 * pat[i % m] + Z[i + 5])))
    fit(`start-${m}x${n}`, t, v, n, m, m + 1, [6, 7])
  }

  // the reading of the timeline and values, with the parameters to fix them
  const base = steps(13).map((x, i) => round2(3 * x + Z[i]))
  const read = (group: string, t: (number | null)[], y: (number | string | null)[], last: number, opts = '') => {
    const [rt, rv] = series(t, y)
    add(group, STAT, ...[1, 2, 6, 7, 8].map((k) => `${STAT}(${rv},${rt},${k},0${opts})`))
    add(group, ETS, ...[1, 2].map((h) => `${ETS}(${last + h},${rv},${rt},0${opts})`))
  }
  const drop = (gone: number[]) => steps(13).filter((x) => !gone.includes(x))
  for (const [name, gone] of Object.entries({ one: [5], two: [5, 6], three: [5, 6, 7], ends: [2, 12], spread: [3, 7, 11] }))
    for (const c of [1, 0]) read(`gap-${name}-${c}`, drop(gone), drop(gone).map((x) => base[x - 1]), 13, `,${c}`)
  // 30 % of what: three missing among 8 given, four among 11 (under 30 % of the full timeline, over 30 % of the points given)
  for (const [name, t] of Object.entries({ m3of8: [1, 2, 4, 5, 7, 8, 10, 11], m4of11: [1, 2, 3, 5, 6, 7, 9, 10, 11, 13, 15] }))
    read(`limit-${name}`, t, t.map((x) => round2(3 * x + Z[x])), t[t.length - 1])
  // a blank value at the start, at the end, two at the end
  for (const [name, at] of Object.entries({ first: [0], last: [12], last2: [11, 12], second: [1] }))
    for (const c of [1, 0]) read(`blank-${name}-${c}`, steps(13), base.map((y, i) => (at.includes(i) ? null : y)), 13, `,${c}`)
  // points at one time by each aggregation
  {
    const t = [1, 2, 2, 3, 4, 5, 5, 5, 6, 7, 8, 9, 10, 10]
    const y = [3, 5, 6, 9, 11, 15, 14, 19, 17, 21, 24, 26, 30, 28]
    for (const a of steps(7)) read(`agg-${a}`, t, y, 10, `,1,${a}`)
  }
  // timelines off the step: halves, a point between two, steps of one and two
  read('off-half', [1, 2, 3.5, 4.5, 5.5, 6.5, 7.5, 8.5, 9.5, 10.5], base.slice(0, 10), 10.5)
  read('off-third', [1, 2, 3.4, 4, 5, 6, 7, 8, 9, 10], base.slice(0, 10), 10)
  read('off-twothirds', [1, 2, 2.6, 4, 5, 6, 7, 8, 9, 10], base.slice(0, 10), 10)
  read('off-mixed', [1, 2, 4, 6, 8, 10, 12, 13, 14, 16], base.slice(0, 10), 16)
  // months: the first, the fifteenth, the last day; quarters
  const day = (y: number, m: number, d: number) => Date.UTC(y, m, d) / 86400000 + 25569
  const months = (d: number, every = 1) => steps(12, 0).map((k) => day(2021, k * every, d))
  for (const [name, t] of Object.entries({ first: months(1), fifteenth: months(15), last: steps(12, 1).map((k) => day(2021, k, 0)), quarters: months(1, 3) })) {
    const [rt, rv] = series(t, base.slice(0, 12))
    add(`month-${name}`, STAT, `${STAT}(${rv},${rt},8,0)`, `${STAT}(${rv},${rt},1,0)`, `${STAT}(${rv},${rt},7,0)`)
    const next = name === 'last' ? day(2022, 1, 0) : name === 'quarters' ? day(2024, 0, 1) : day(2022, 0, name === 'first' ? 1 : 15)
    add(`month-${name}`, ETS, `${ETS}(${next},${rv},${rt},0)`, `${ETS}(${next + 15},${rv},${rt},0)`)
  }

  // intervals between steps and far ahead
  {
    const [t, v] = series(steps(10), [3, 5, 9, 11, 15, 17, 21, 24, 26, 30])
    for (const x of [10.25, 10.75, 11.5, 12.5, 13.75, 20]) add('interval', CI, `${CI}(${x},${v},${t},0.95,0)`)
    for (const x of [10.25, 11.5, 12.5]) add('interval', ETS, `${ETS}(${x},${v},${t},0)`)
    const [ts, vs] = series(steps(12), steps(12).map((x, i) => round2(20 + 0.5 * x + PATTERN[i % 4] + Z[i])))
    for (const x of [12.5, 13.5, 15.5, 16, 17, 21]) add('interval', CI, `${CI}(${x},${vs},${ts},0.95,4)`)
    for (const x of [12.5, 13.5, 16.5, 12, 11, 10.5]) add('interval', ETS, `${ETS}(${x},${vs},${ts},4)`)
  }
  // the longest season taken: a season of n − 1, n, n + 1 points and far beyond
  {
    const [t, v] = series(steps(10), [3, 5, 9, 11, 15, 17, 21, 24, 26, 30])
    for (const m of [7, 8, 9, 10, 11, 100, 8760, 8761, 10000, 1e6]) add('longest', STAT, `${STAT}(${v},${t},3,${m})`)
    for (const m of [8, 9, 10, 11]) add('longest', ETS, `${ETS}(11,${v},${t},${m})`)
  }
  // seasons to detect: one pattern at every length, a sine, steps, noise of several seeds
  for (const p of [2, 3, 4, 5, 6, 7, 8, 10, 12]) {
    const pat = steps(p).map((j) => round2(5 * Math.sin((2 * Math.PI * j) / p)))
    for (const [k, n] of [3 * p, 5 * p].entries()) {
      const [t, v] = series(steps(n), steps(n).map((x, i) => round2(20 + 0.2 * x + pat[i % p] + 0.5 * Z[i % 40])))
      add(`detect-${p}-${k}`, SEA, `${SEA}(${v},${t})`)
      add(`detect-${p}-${k}`, STAT, `${STAT}(${v},${t},1)`, `${STAT}(${v},${t},3)`)
    }
  }
  for (let k = 0; k < 6; k++) {
    const [t, v] = series(steps(24), noise(300 + k, 24).map((x) => round2(10 + 4 * x)))
    add(`detect-noise-${k}`, SEA, `${SEA}(${v},${t})`)
  }
  return out
}

const PROBES = build()

export const SHEETS: SheetCells = { E: cells }

export const buildProbes = (): Probe[] => PROBES

/** The model's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./ets-fit-2.predicted.json', import.meta.url), 'utf8'))
