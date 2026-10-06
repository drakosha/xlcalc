/**
 * Statistical functions.
 *
 * AVEDEV … PROB are ported from formula.js 4.6.1 (https://github.com/formulajs/formulajs),
 * src/statistical.js (BINOMDIST and NORMDIST as exported by src/compatibility.js): Copyright (c) 2014
 * Sutoiku, Inc., MIT License (full text in THIRD_PARTY.md); the distributions they call come from jStat
 * (`./jstat.ts`). Changed: arguments are coerced by the engine's argument layer (`src/args.ts`), empty input
 * by its rule 3; COUNTA counts empty text (""), as Excel does; BINOMDIST, BINOM.DIST.RANGE, F.DIST.RT,
 * GAMMALN and PROB check their arguments' domain as Excel documents it (#NUM!), F.DIST.RT truncates the
 * degrees of freedom; PROB pairs each value with its own probability (formula.js sorted the values in place
 * and lost the pairing).
 * COUNTIFS is this engine's own.
 */
import { eachValue, gridAt, type Cx, type FnDef, type Grid, type Segment } from '../args.js'
import { E, isErr, type V } from '../values.js'
import { combin } from './common.js'
import { binomialCdf, binomialPdf, centralFCdf, gammaln, normalCdf, normalPdf } from './jstat.js'

const SQRT2PI = 2.5066282746310002

const sum = (xs: number[]) => xs.reduce((s, x) => s + x, 0)
// a loop, not Math.max(...xs): a whole column would overflow the call stack
const max = (xs: number[]) => xs.reduce((m, x) => (x > m ? x : m), -Infinity)
const min = (xs: number[]) => xs.reduce((m, x) => (x < m ? x : m), Infinity)

function median(xs: number[]): number {
  const s = [...xs].sort((a, b) => a - b)
  const m = s.length >> 1
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2
}

function normDist(x: number, mean: number, sd: number, cumulative: boolean): V {
  if (sd <= 0) return E.NUM
  return cumulative ? normalCdf(x, mean, sd) : normalPdf(x, mean, sd)
}

function binomDist(k: number, n: number, p: number, cumulative: boolean): V {
  if (k < 0 || k > n || p < 0 || p > 1) return E.NUM
  return cumulative ? binomialCdf(k, n, p) : binomialPdf(k, n, p)
}

function gammaLn(x: number): V {
  return x <= 0 ? E.NUM : gammaln(x)
}

export const functions: Record<string, FnDef> = {
  AVEDEV: {
    args: 'N',
    rep: 1,
    empty: 'NUM',
    fn: (xs: number[]) => {
      const mean = sum(xs) / xs.length
      return sum(xs.map((x) => Math.abs(x - mean))) / xs.length
    },
  },
  AVERAGE: { args: 'N', rep: 1, empty: 'DIV0', fn: (xs: number[]) => sum(xs) / xs.length },
  'BINOM.DIST.RANGE': {
    args: 'i n i i?',
    fn: (n: number, p: number, s: number, s2 = s) => {
      if (n < 0 || p < 0 || p > 1 || s < 0 || s > n || s2 < s || s2 > n) return E.NUM
      let r = 0
      for (let k = s; k <= s2; k++) r += combin(n, k) * Math.pow(p, k) * Math.pow(1 - p, n - k)
      return r
    },
  },
  BINOMDIST: { args: 'i i n b', fn: binomDist },
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
  COUNTIFS: {
    args: 'R a',
    rep: 2,
    cx: true,
    fn: (cx: Cx, rest: V[]) => {
      const pairs: [V, V][] = []
      for (let i = 0; i < rest.length; i += 2) pairs.push([rest[i], rest[i + 1]])
      const r = cx.select(pairs)
      return isErr(r) ? r : r.sel.length + r.rest
    },
  },
  'F.DIST.RT': {
    args: 'n n n',
    fn: (x: number, d1: number, d2: number) => {
      d1 = Math.trunc(d1)
      d2 = Math.trunc(d2)
      return x < 0 || d1 < 1 || d2 < 1 ? E.NUM : 1 - centralFCdf(x, d1, d2)
    },
  },
  FISHER: { args: 'n', fn: (x: number) => Math.log((1 + x) / (1 - x)) / 2 },
  GAMMALN: { args: 'n', fn: gammaLn },
  'GAMMALN.PRECISE': { args: 'n', fn: gammaLn },
  GAUSS: { args: 'n', fn: (z: number) => normalCdf(z, 0, 1) - 0.5 },
  MAX: { args: 'N', rep: 1, fn: (xs: number[]) => (xs.length ? max(xs) : 0) },
  MEDIAN: { args: 'N', rep: 1, empty: 'NUM', fn: median },
  MIN: { args: 'N', rep: 1, fn: (xs: number[]) => (xs.length ? min(xs) : 0) },
  'NORM.DIST': { args: 'n n n b', fn: normDist },
  'NORM.S.DIST': { args: 'n b', fn: (z: number, cumulative: boolean) => (cumulative ? normalCdf(z, 0, 1) : normalPdf(z, 0, 1)) },
  NORMDIST: { args: 'n n n b', fn: normDist },
  PHI: { args: 'n', fn: (x: number) => Math.exp(-0.5 * x * x) / SQRT2PI },
  PROB: {
    args: 'M M n n?',
    fn: (xs: Grid, ps: Grid, lo: number, hi = lo) => {
      if (xs.rows * xs.cols !== ps.rows * ps.cols) return E.NA
      let total = 0
      let r = 0
      for (let i = 0, n = Math.max(xs.h * xs.cols, ps.h * ps.cols); i < n; i++) {
        const x = gridAt(xs, i)
        const p = gridAt(ps, i)
        if (x === null && p === null) continue
        if (typeof x !== 'number' || typeof p !== 'number') return E.VALUE
        if (p <= 0 || p > 1) return E.NUM
        total += p
        if (x >= lo && x <= hi) r += p
      }
      return Number(total.toPrecision(15)) === 1 ? r : E.NUM
    },
  },
}
