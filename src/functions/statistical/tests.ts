/**
 * Hypothesis tests. The statistics follow formula.js 4.6.1 (https://github.com/formulajs/formulajs),
 * src/statistical.js: Copyright (c) 2014 Sutoiku, Inc., MIT License (THIRD_PARTY.md); the p-values,
 * tails, test types and errors follow Excel (formula.js returned the F statistic for F.TEST, ignored the
 * tails and type of T.TEST and rounded CHISQ.TEST to six digits).
 */
import { gammaPQ, normCdf } from './special.js'
import { fCdf, tCdf } from './distributions.js'
import { gridAt, type FnDef, type Grid } from '../../args.js'
import { E, isErr, type V } from '../../values.js'
import { devsq, gridError, gridNumbers, mean, numberPairs } from './common.js'

function chisqTest(actual: Grid, expected: Grid): V {
  const err = gridError(actual) ?? gridError(expected)
  if (err) return err
  const n = actual.rows * actual.cols
  if (n !== expected.rows * expected.cols || n === 1) return E.NA
  let x2 = 0
  let pairs = 0
  for (let i = 0; i < n; i++) {
    const a = gridAt(actual, i)
    const e = gridAt(expected, i)
    if (typeof a !== 'number' || typeof e !== 'number') continue
    if (e === 0) return E.DIV0
    x2 += ((a - e) * (a - e)) / e
    pairs++
  }
  if (!pairs) return E.DIV0
  const df = actual.rows > 1 && actual.cols > 1 ? (actual.rows - 1) * (actual.cols - 1) : n - 1
  return gammaPQ(df / 2, x2 / 2)[1]
}

function fTest(a: number[], b: number[]): V {
  if (a.length < 2 || b.length < 2) return E.DIV0
  const va = devsq(a) / (a.length - 1)
  const vb = devsq(b) / (b.length - 1)
  if (va === 0 || vb === 0) return E.DIV0
  const up = fCdf(va / vb, a.length - 1, b.length - 1)[1]
  return 2 * Math.min(up, 1 - up)
}

function tTest(ga: Grid, gb: Grid, tails: number, type: number): V {
  if ((tails !== 1 && tails !== 2) || type < 1 || type > 3) return E.NUM
  let t: number
  let df: number
  if (type === 1) {
    const p = numberPairs(ga, gb)
    if (isErr(p)) return p
    const d = p[0].map((x, i) => x - p[1][i])
    const n = d.length
    const sd = n < 2 ? 0 : Math.sqrt(devsq(d) / (n - 1))
    if (sd === 0) return E.DIV0
    t = mean(d) / (sd / Math.sqrt(n))
    df = n - 1
  } else {
    const a = gridNumbers(ga)
    if (isErr(a)) return a
    const b = gridNumbers(gb)
    if (isErr(b)) return b
    const na = a.length
    const nb = b.length
    if (na < 2 || nb < 2) return E.DIV0
    const va = devsq(a) / (na - 1)
    const vb = devsq(b) / (nb - 1)
    if (type === 2) {
      const pooled = (devsq(a) + devsq(b)) / (na + nb - 2)
      if (pooled === 0) return E.DIV0
      t = (mean(a) - mean(b)) / Math.sqrt(pooled * (1 / na + 1 / nb))
      df = na + nb - 2
    } else {
      const se2 = va / na + vb / nb
      if (se2 === 0) return E.DIV0
      t = (mean(a) - mean(b)) / Math.sqrt(se2)
      // Welch–Satterthwaite degrees of freedom
      df = (se2 * se2) / ((va / na) ** 2 / (na - 1) + (vb / nb) ** 2 / (nb - 1))
    }
  }
  return tails * tCdf(Math.abs(t), df)[1]
}

function zTest(xs: number[], x: number, sigma?: number): V {
  const n = xs.length
  if (!n) return E.NA
  if (sigma !== undefined && sigma <= 0) return E.NUM
  if (n < 2) return E.DIV0
  const s = sigma ?? Math.sqrt(devsq(xs) / (n - 1))
  if (s === 0) return E.DIV0
  return normCdf((mean(xs) - x) / (s / Math.sqrt(n)))[1]
}

export const functions: Record<string, FnDef> = {
  'CHISQ.TEST': { args: 'G G', fn: chisqTest },
  'F.TEST': { args: 'N N', fn: fTest },
  'T.TEST': { args: 'G G i i', fn: tTest },
  'Z.TEST': { args: 'N n n?', fn: zTest },
}
