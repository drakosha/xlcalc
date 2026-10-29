/**
 * Depreciation. DB, DDB, SLN and SYD follow formula.js 4.6.1 (src/financial.js), Copyright (c) 2014
 * Sutoiku, Inc., MIT; fixed against Excel. AMORDEGRC and AMORLINC are this engine's own.
 */
import { E } from '../../values.js'
import type { FnDef } from '../../args.js'
import { spend } from '../../exec/budget.js'
import { roundDigits } from '../common.js'
import { atp, need } from '../date/atp.js'
import { isLeap, ymd } from '../../dates.js'
import { validBasis, yearFrac } from '../date/daycount.js'

const round = (x: number, digits: number) => roundDigits(x, digits, Math.round)
// roundDigits to 0 digits without its decimal text: AMORDEGRC rounds once per period
const round0 = (x: number) => Math.sign(x) * Math.round(Math.abs(x))

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
/**
 * VDB from `start` to `end`. Excel's periods run from the fractional part of `start` (0.5, 1.5, 2.5… for a start of 8.5),
 * after a first one of that length: each period declines the book value by the rate, never below the salvage (with a
 * rate of 1 or more, all of it but the salvage), and from the period where the rest spread evenly over the rest of the
 * life gives more, by that even amount; a period cut short takes its share of its amount. Without the switch, each whole
 * period's double-declining amount, the first and last ones in part.
 */
function vdb(cost: number, salvage: number, life: number, start: number, end: number, factor: number, noSwitch: boolean): number {
  if (noSwitch) {
    const s0 = Math.floor(start)
    const e0 = Math.ceil(end)
    let sum = 0
    for (let i = s0 + 1; i <= e0; i++) {
      spend(1)
      const d = ddb(cost, salvage, life, i, factor)
      sum += i === s0 + 1 ? d * (Math.min(end, s0 + 1) - start) : i === e0 ? d * (end + 1 - e0) : d
    }
    return sum
  }
  const rate = factor / life
  const frac = start - Math.floor(start)
  // period k runs from `at(k)` to `at(k + 1)`: 0, then the fractional part of start and whole periods after it
  const at = (k: number) => (k === 0 ? 0 : frac ? frac + k - 1 : k)
  const first = frac ? Math.floor(start) + 1 : start
  let left = cost - salvage
  let even = 0
  let sum = 0
  for (let k = 0; at(k) < end; k++) {
    spend(1)
    const t = at(k)
    let amount = even
    if (!even) {
      // the decline before the salvage limits it decides the switch; a rate of 1 or more takes everything at once
      const declined = rate >= 1 ? left : (left + salvage) * rate
      if (rate < 1 && left / (life - t) > declined) even = amount = left / (life - t)
      else amount = Math.min(declined, left)
    }
    const part = amount * (Math.min(at(k + 1), end) - t)
    left -= part
    if (k >= first) sum += part
  }
  return sum
}

const coefficient = (life: number) => (life <= 2 ? NaN : life <= 4 ? 1.5 : life <= 6 ? 2 : 2.5)

/** AMORDEGRC and AMORLINC check cost, dates, salvage, period and rate before they read the basis. */
const ASSET = { check: { 5: need((rate: number, [cost, purchased, first, salvage, period]: readonly number[]) => rate > 0 && cost >= 0 && salvage >= 0 && salvage <= cost && period >= 0 && purchased <= first) } }
const badBasis = (basis: number) => !validBasis(basis) || basis === 2

export const functions: Record<string, FnDef> = {
  // declining by rate × coefficient, rounded; the first period no more than cost less salvage; the year before the life's
  // end takes half the rest, the last year the same again; the period that passes the salvage takes its whole amount, later ones nothing
  AMORDEGRC: atp('x e e x x x j?', (cost: number, purchased: number, first: number, salvage: number, period: number, rate: number, basis = 0) => {
    if (badBasis(basis)) return E.NUM
    const r = rate * coefficient(1 / rate)
    if (Number.isNaN(r)) return E.NUM
    const last = Math.floor(1 / rate)
    let dep = Math.min(round0(amorFirst(purchased, first, basis) * r * cost), cost - salvage)
    let book = cost - dep
    let rest = book - salvage
    const n = Math.trunc(period)
    // past the life's end nothing is left
    if (n > last) return 0
    for (let k = 0; k < n; k++) {
      spend(1)
      // once the rounded amount is 0 nothing changes until the last two years
      if (dep === 0 && k > 0 && k + 2 < last) k = Math.min(n, last - 2) - 1
      if (k + 1 === last) {
        book = 0
        continue
      }
      dep = round0(k + 2 === last ? book * 0.5 : r * book)
      rest -= dep
      if (rest < 0) return n - k <= 1 ? dep : 0
      book -= dep
    }
    return dep
  }, ASSET),
  AMORLINC: atp('x e e x x x j?', (cost: number, purchased: number, first: number, salvage: number, period: number, rate: number, basis = 0) => {
    if (badBasis(basis)) return E.NUM
    const full = cost * rate
    // the first period no more than cost less salvage
    const first0 = Math.min(amorFirst(purchased, first, basis) * full, cost - salvage)
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
  VDB: {
    args: 'n n n n n n? b?',
    fn: (cost: number, salvage: number, life: number, start: number, end: number, factor = 2, noSwitch = false) =>
      cost < 0 || salvage < 0 || life <= 0 || start < 0 || end < start || end > life || factor < 0 ? E.NUM : vdb(cost, salvage, life, start, end, factor, noSwitch),
  },
  SLN: { args: 'n n n', fn: (cost: number, salvage: number, life: number) => (life === 0 ? E.DIV0 : (cost - salvage) / life) },
  SYD: {
    args: 'n n n n',
    fn: (cost: number, salvage: number, life: number, per: number) =>
      salvage < 0 || life <= 0 || per <= 0 || per > life ? E.NUM : ((cost - salvage) * (life - per + 1) * 2) / (life * (life + 1)),
  },
}
