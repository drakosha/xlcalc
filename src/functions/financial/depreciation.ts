/**
 * Depreciation. DB, DDB, SLN and SYD follow formula.js 4.6.1 (src/financial.js), Copyright (c) 2014
 * Sutoiku, Inc., MIT; fixed against Excel. AMORDEGRC and AMORLINC are this engine's own.
 */
import { E } from '../../values.js'
import type { FnDef } from '../../args.js'
import { roundDigits } from '../common.js'
import { atp, need } from '../date/atp.js'
import { isLeap, ymd } from '../../dates.js'
import { validBasis, yearFrac } from '../date/daycount.js'

const round = (x: number, digits: number) => roundDigits(x, digits, Math.round)

/** Double-declining balance for a (possibly fractional) period, never below salvage. */
function ddb(cost: number, salvage: number, life: number, period: number, factor: number): number {
  let rate = factor / life
  let before: number
  if (rate >= 1) {
    rate = 1
    before = period === 1 ? cost : 0
  } else before = cost * Math.pow(1 - rate, period - 1)
  const after = cost * Math.pow(1 - rate, period)
  return Math.max(0, after < salvage ? before - salvage : before - after)
}

/** The French methods' first period: the year fraction to the first period's end (on basis 1 in the purchase year's days); none at all counts as a full year. */
function amorFirst(purchased: number, first: number, basis: number): number {
  const f = basis === 1 ? (first - purchased) / (isLeap(ymd(purchased).y) ? 366 : 365) : yearFrac(purchased, first, basis)
  return f === 0 ? 1 : f
}

/** AMORDEGRC's coefficient by asset life 1/rate; a life of 2 years or less is #NUM!. */
const coefficient = (life: number) => (life <= 2 ? NaN : life <= 4 ? 1.5 : life <= 6 ? 2 : 2.5)

/** AMORDEGRC and AMORLINC check cost, dates, salvage, period and rate before they read the basis. */
const ASSET = { check: { 5: need((rate: number, [cost, purchased, first, salvage, period]: readonly number[]) => rate > 0 && cost >= 0 && salvage >= 0 && salvage <= cost && period >= 0 && purchased <= first) } }
const badBasis = (basis: number) => !validBasis(basis) || basis === 2

export const functions: Record<string, FnDef> = {
  // declining by rate × coefficient, rounded; the year before the life's end takes half the rest, the last year the same again
  AMORDEGRC: atp('x e e x x x j?', (cost: number, purchased: number, first: number, salvage: number, period: number, rate: number, basis = 0) => {
    if (badBasis(basis)) return E.NUM
    const r = rate * coefficient(1 / rate)
    if (Number.isNaN(r)) return E.NUM
    const last = Math.floor(1 / rate)
    let dep = round(amorFirst(purchased, first, basis) * r * cost, 0)
    let book = cost - dep
    let rest = book - salvage
    const n = Math.trunc(period)
    // past the life's end nothing is left
    if (n > last) return 0
    for (let k = 0; k < n; k++) {
      // once the rounded amount is 0 nothing changes until the last two years
      if (dep === 0 && k > 0 && k + 2 < last) k = Math.min(n, last - 2) - 1
      if (k + 1 === last) {
        book = 0
        continue
      }
      dep = round(k + 2 === last ? book * 0.5 : r * book, 0)
      rest -= dep
      if (rest < 0) return n - k <= 1 ? round(book * 0.5, 0) : 0
      book -= dep
    }
    return dep
  }, ASSET),
  AMORLINC: atp('x e e x x x j?', (cost: number, purchased: number, first: number, salvage: number, period: number, rate: number, basis = 0) => {
    if (badBasis(basis)) return E.NUM
    const full = cost * rate
    const first0 = amorFirst(purchased, first, basis) * full
    const n = Math.trunc(period)
    if (n === 0) return first0
    const fulls = Math.floor((cost - salvage - first0) / full)
    if (n <= fulls) return full
    if (n === fulls + 1) return cost - salvage - full * fulls - first0
    return 0
  }, ASSET),
  // the rate rounded to three decimals; a partial first year of `month` months and its remainder after `life`
  DB: {
    args: 'n n n n w?',
    fn: (cost: number, salvage: number, life: number, period: number, month = 12) => {
      if (month < 1 || month > 12 || life <= 0 || period <= 0 || period > life + (month < 12 ? 1 : 0)) return E.NUM
      if (cost === 0) return 0
      const rate = round(1 - Math.pow(salvage / cost, 1 / life), 3)
      const first = (cost * rate * month) / 12
      const p = Math.trunc(period)
      if (p <= 1) return first
      // the book value declines geometrically after the first year
      const dep = (cost - first) * Math.pow(1 - rate, p - 2) * rate
      return p > life ? (dep * (12 - month)) / 12 : dep
    },
  },
  DDB: {
    args: 'n n n n n?',
    fn: (cost: number, salvage: number, life: number, period: number, factor = 2) =>
      cost < 0 || salvage < 0 || life <= 0 || period <= 0 || period > life || factor <= 0 ? E.NUM : ddb(cost, salvage, life, Math.max(1, period), factor),
  },
  SLN: { args: 'n n n', fn: (cost: number, salvage: number, life: number) => (life === 0 ? E.DIV0 : (cost - salvage) / life) },
  SYD: {
    args: 'n n n n',
    fn: (cost: number, salvage: number, life: number, per: number) =>
      salvage < 0 || life <= 0 || per <= 0 || per > life ? E.NUM : ((cost - salvage) * (life - per + 1) * 2) / (life * (life + 1)),
  },
}
