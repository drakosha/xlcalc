/**
 * Descriptive statistics. AVEDEV, AVERAGE, COUNTA, COUNTBLANK, MAX, MEDIAN, MIN and the formulas of the rest
 * come from formula.js 4.6.1 (https://github.com/formulajs/formulajs), src/statistical.js: Copyright (c) 2014
 * Sutoiku, Inc., MIT License (THIRD_PARTY.md); empty input, coercion and domain checks follow Excel.
 */
import { eachValue, type Arg, type Cx, type FnDef, type Grid, type Segment } from '../../args.js'
import { Arr } from '../../arrays/value.js'
import { spend } from '../../exec/budget.js'
import type { Rules } from '../../locale/rules.js'
import { E, isErr, type CErr, type V } from '../../values.js'
import { max, min, NEWER_MAX_ARGS, sum, sumZeroed } from '../common.js'
import { functions as trig } from '../math/trig.js'
import { areaNumbers, devsq, interpolate, mean, sorted, stdev, variance } from './common.js'

function median(xs: number[]): number {
  const s = sorted(xs)
  const m = s.length >> 1
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2
}

/** AVERAGE: 0 where SUM's last addition leaves a remainder Excel makes 0 (the set `sum-zero-held`), else the stable mean. */
const average = (xs: number[]) => (sumZeroed(xs) ? 0 : mean(xs))

/** Σ((x − mean)/s)^k with s the sample (ddof 1) or population (ddof 0) deviation; #DIV/0! when s is 0. */
function moment(xs: number[], k: number, ddof: number): number | undefined {
  const m = mean(xs)
  const s = Math.sqrt(devsq(xs) / (xs.length - ddof))
  if (!(s > 0)) return undefined
  let acc = 0
  for (const x of xs) acc += ((x - m) / s) ** k
  return acc
}

function percentileInc(xs: number[], k: number): V {
  if (k < 0 || k > 1) return E.NUM
  return interpolate(sorted(xs), k * (xs.length - 1))
}

function percentileExc(xs: number[], k: number): V {
  const n = xs.length
  // Excel compares k with 1/(n+1) and n/(n+1), not the rank with 1 and n
  if (k < 1 / (n + 1) || k > n / (n + 1)) return E.NUM
  return interpolate(sorted(xs), Math.max(k * (n + 1) - 1, 0))
}

// the quart is checked before it is truncated: QUARTILE(x, -0.5) is #NUM!, QUARTILE(x, 4.9) the maximum
const quartileInc = (xs: number[], q: number): V => (q < 0 || q >= 5 ? E.NUM : percentileInc(xs, Math.trunc(q) / 4))
const quartileExc = (xs: number[], q: number): V => (q < 1 || q >= 4 ? E.NUM : percentileExc(xs, Math.trunc(q) / 4))

/** PERCENTRANK: the rank of x among the data, interpolated, truncated (not rounded) to `digits`. */
function percentRank(xs: number[], x: number, digits = 3, exclusive: boolean): V {
  if (digits < 1) return E.NUM
  const s = sorted(xs)
  const n = s.length
  if (x < s[0] || x > s[n - 1]) return E.NA
  if (n === 1) return 1
  let i = 0
  while (s[i] < x) i++
  // position among the n values: the first equal one, or between s[i−1] and s[i]
  const pos = s[i] === x ? i : i - 1 + (x - s[i - 1]) / (s[i] - s[i - 1])
  const r = exclusive ? (pos + 1) / (n + 1) : pos / (n - 1)
  const p = 10 ** Math.min(digits, 15)
  // the product read at 15 digits, as Excel sees it: 0.285 * 1000 is 285
  return Math.floor(Number((r * p).toPrecision(15))) / p
}

const percentRankInc = (xs: number[], x: number, digits?: number) => percentRank(xs, x, digits, false)
const percentRankExc = (xs: number[], x: number, digits?: number) => percentRank(xs, x, digits, true)

/** RANK's order: a number or a boolean; text, even "0", is #VALUE!. */
function descending(order: V): boolean | V {
  if (typeof order === 'string') return E.VALUE
  return !order
}

function rank(x: number, xs: number[], order: V | undefined, average: boolean): V {
  const desc = descending(order ?? 0)
  if (typeof desc !== 'boolean') return desc
  let before = 0
  let equal = 0
  for (const v of xs) {
    if (v === x) equal++
    else if (desc ? v > x : v < x) before++
  }
  if (!equal) return E.NA
  return average ? before + (equal + 1) / 2 : before + 1
}

const rankEq = (x: number, xs: number[], order?: V) => rank(x, xs, order, false)
const rankAvg = (x: number, xs: number[], order?: V) => rank(x, xs, order, true)

/**
 * MODE: the most frequent numbers (seen more than once), in the order they first occur; text and booleans are
 * skipped even as direct arguments, but a reference to one blank cell is #VALUE!; none is #N/A.
 */
function modes(grids: Grid[]): number[] | CErr {
  const counts = new Map<number, number>()
  let top = 1
  for (const g of grids) {
    if (g.rows * g.cols === 1 && (g.v[0] ?? null) === null) return E.VALUE
    for (const v of g.v) {
      if (isErr(v)) return v
      if (typeof v !== 'number') continue
      const c = (counts.get(v) ?? 0) + 1
      counts.set(v, c)
      if (c > top) top = c
    }
  }
  if (top === 1) return E.NA
  const out = new Set<number>()
  for (const g of grids) for (const v of g.v) if (typeof v === 'number' && counts.get(v) === top) out.add(v)
  return [...out]
}

function mode(grids: Grid[]): V {
  const m = modes(grids)
  return isErr(m) ? m : m[0]
}

/**
 * FREQUENCY: how many numbers of the data fall at or below each bin and above the bins before it (bins taken
 * in sorted order, answered in their own order; a repeated bin counts nothing), and above the last bin. Bins
 * without numbers are one bin at 0, as in Excel. A union or a reference across sheets gives the numbers of each area.
 */
function frequency(cx: Cx, data: Arg, bins: Arg): V | Arr {
  const xs = areaNumbers(cx, data)
  if (isErr(xs)) return xs
  const found = areaNumbers(cx, bins)
  if (isErr(found)) return found
  const bs = found.length ? found : [0]
  const order = bs.map((_, i) => i).sort((a, b) => bs[a] - bs[b] || a - b)
  const counts = new Array<number>(bs.length + 1).fill(0)
  spend(xs.length)
  for (const x of xs) {
    // the first bin in sorted order at or above x, by halves: bs[order[p]] never decreases in p
    let lo = 0
    let hi = order.length
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (x <= bs[order[mid]]) hi = mid
      else lo = mid + 1
    }
    counts[lo < order.length ? order[lo] : bs.length]++
  }
  return new Arr(counts.length, 1, counts.length, 1, counts)
}

/** COUNT: numbers in ranges; a direct argument counts if it reads as a number (booleans, numeric and date text, a blank: `COUNT(1,,2)` is 3). */
function countable(v: V, inRange: boolean, rules: Rules): boolean {
  if (typeof v === 'number') return true
  if (inRange || isErr(v)) return false
  if (v === null || typeof v === 'boolean') return true
  return !isErr(rules.toNum(v)) || rules.dateText(v) !== null
}

export const functions: Record<string, FnDef> = {
  AVEDEV: {
    args: 'N',
    rep: 1,
    empty: 'NUM',
    fn: (xs: number[]) => {
      const m = mean(xs)
      return sum(xs.map((x) => Math.abs(x - m))) / xs.length
    },
  },
  AVERAGE: { args: 'N', rep: 1, empty: 'DIV0', fn: average },
  AVERAGEA: { args: 'Y', rep: 1, empty: 'DIV0', fn: average },
  COUNT: {
    args: 'V',
    rep: 1,
    cx: true,
    fn: (cx: Cx, segs: Segment[]) => {
      let n = 0
      eachValue(segs, (v, inRange) => {
        if (countable(v, inRange, cx.rules)) n++
      })
      return n
    },
  },
  // a range counts its non-blank cells (errors and "" included); every direct argument counts
  COUNTA: {
    args: 'V',
    rep: 1,
    fn: (segs: Segment[]) => {
      let n = 0
      eachValue(segs, (v, inRange) => {
        if (!inRange || v !== null) n++
      })
      return n
    },
  },
  COUNTBLANK: {
    args: 'G',
    fn: (g: Grid) => {
      let filled = 0
      for (const v of g.v) if (v !== null && v !== '') filled++
      return g.rows * g.cols - filled
    },
  },
  DEVSQ: { args: 'N', rep: 1, empty: 'NUM', fn: devsq },
  // the inverse hyperbolic tangent and the tangent itself
  FISHER: trig.ATANH,
  FISHERINV: trig.TANH,
  // through logarithms: the product of the data would overflow
  GEOMEAN: { args: 'N', rep: 1, empty: 'NUM', fn: (xs: number[]) => (xs.some((x) => x <= 0) ? E.NUM : Math.exp(mean(xs.map(Math.log)))) },
  HARMEAN: { args: 'N', rep: 1, empty: 'NA', fn: (xs: number[]) => (xs.some((x) => x <= 0) ? E.NUM : xs.length / sum(xs.map((x) => 1 / x))) },
  KURT: {
    args: 'N',
    rep: 1,
    fn: (xs: number[]) => {
      const n = xs.length
      const m4 = n < 4 ? undefined : moment(xs, 4, 1)
      if (m4 === undefined) return E.DIV0
      return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * m4 - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3))
    },
  },
  LARGE: {
    args: 'N n',
    empty: 'NUM',
    fn: (xs: number[], k: number) => (k < 1 || k > xs.length ? E.NUM : sorted(xs)[xs.length - Math.ceil(k)]),
  },
  MAX: { args: 'N', rep: 1, fn: max },
  MAXA: { args: 'A', rep: 1, fn: max },
  MEDIAN: { args: 'N', rep: 1, empty: 'NUM', fn: median },
  MIN: { args: 'N', rep: 1, fn: min },
  MINA: { args: 'A', rep: 1, fn: min },
  'MODE.MULT': {
    args: 'G',
    rep: 1,
    max: NEWER_MAX_ARGS,
    array: true,
    fn: (grids: Grid[]) => {
      const m = modes(grids)
      return isErr(m) ? m : new Arr(m.length, 1, m.length, 1, m)
    },
  },
  'MODE.SNGL': { args: 'G', rep: 1, max: NEWER_MAX_ARGS, fn: mode },
  FREQUENCY: { args: 'R R', array: true, cx: true, areas: true, fn: frequency },
  'PERCENTILE.EXC': { args: 'N n', empty: 'NUM', fn: percentileExc },
  'PERCENTILE.INC': { args: 'N n', empty: 'NUM', fn: percentileInc },
  'PERCENTRANK.EXC': { args: 'N n i?', empty: 'NA', fn: percentRankExc },
  'PERCENTRANK.INC': { args: 'N n i?', empty: 'NA', fn: percentRankInc },
  'QUARTILE.EXC': { args: 'N n', empty: 'NUM', fn: quartileExc },
  'QUARTILE.INC': { args: 'N n', empty: 'NUM', fn: quartileInc },
  'RANK.AVG': { args: 'n N v?', fn: rankAvg },
  'RANK.EQ': { args: 'n N v?', fn: rankEq },
  SKEW: {
    args: 'N',
    rep: 1,
    fn: (xs: number[]) => {
      const n = xs.length
      const m3 = n < 3 ? undefined : moment(xs, 3, 1)
      return m3 === undefined ? E.DIV0 : (n / ((n - 1) * (n - 2))) * m3
    },
  },
  'SKEW.P': {
    args: 'N',
    rep: 1,
    max: NEWER_MAX_ARGS,
    fn: (xs: number[]) => {
      const m3 = xs.length < 3 ? undefined : moment(xs, 3, 0)
      return m3 === undefined ? E.DIV0 : m3 / xs.length
    },
  },
  SMALL: {
    args: 'N n',
    empty: 'NUM',
    fn: (xs: number[], k: number) => (k < 1 || k > xs.length ? E.NUM : sorted(xs)[Math.floor(k) - 1]),
  },
  STANDARDIZE: { args: 'n n n', fn: (x: number, m: number, sd: number) => (sd <= 0 ? E.NUM : (x - m) / sd) },
  'STDEV.P': { args: 'N', rep: 1, max: NEWER_MAX_ARGS, fn: (xs: number[]) => stdev(xs, 0) },
  'STDEV.S': { args: 'N', rep: 1, max: NEWER_MAX_ARGS, fn: (xs: number[]) => stdev(xs, 1) },
  STDEVA: { args: 'A', rep: 1, fn: (xs: number[]) => stdev(xs, 1) },
  STDEVPA: { args: 'A', rep: 1, fn: (xs: number[]) => stdev(xs, 0) },
  TRIMMEAN: {
    args: 'N n',
    empty: 'NUM',
    fn: (xs: number[], p: number) => {
      if (p < 0 || p >= 1) return E.NUM
      // the number trimmed is rounded down to an even count, half from each end
      const t = Math.floor((xs.length * p) / 2)
      return mean(sorted(xs).slice(t, xs.length - t))
    },
  },
  'VAR.P': { args: 'N', rep: 1, max: NEWER_MAX_ARGS, fn: (xs: number[]) => variance(xs, 0) },
  'VAR.S': { args: 'N', rep: 1, max: NEWER_MAX_ARGS, fn: (xs: number[]) => variance(xs, 1) },
  VARA: { args: 'A', rep: 1, fn: (xs: number[]) => variance(xs, 1) },
  VARPA: { args: 'A', rep: 1, fn: (xs: number[]) => variance(xs, 0) },
}
