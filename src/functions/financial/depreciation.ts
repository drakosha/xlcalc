/**
 * Depreciation. DB, DDB, SLN and SYD follow formula.js 4.6.1 (src/financial.js), Copyright (c) 2014
 * Sutoiku, Inc., MIT; fixed against Excel. VDB, AMORDEGRC and AMORLINC are this engine's own.
 */
import type { FnDef } from '../../args.js'
import { E } from '../../values.js'
import { roundDigits } from '../common.js'
import { atp } from '../date/atp.js'
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

/** VDB's schedule: declining balance, switching to straight line once that is larger (unless `noSwitch`). */
function schedule(cost: number, salvage: number, life: number, factor: number, noSwitch: boolean, periods: number): number[] {
  const out: number[] = []
  let book = cost
  let straight = false
  let sl = 0
  for (let p = 1; p <= periods; p++) {
    if (noSwitch) {
      out.push(ddb(cost, salvage, life, p, factor))
      continue
    }
    if (!straight) {
      const d = Math.min(book * Math.min(factor / life, 1), Math.max(0, book - salvage))
      sl = (book - salvage) / (life - p + 1)
      if (sl > d) straight = true
      else {
        out.push(d)
        book -= d
        continue
      }
    }
    out.push(p <= Math.ceil(life) ? sl : 0)
  }
  return out
}

/** The French linear and degressive methods: the first period prorated by `yearFrac`. */
function amorFirst(cost: number, purchased: number, first: number, rate: number, basis: number): number {
  return yearFrac(purchased, first, basis) * rate * cost
}

export const functions: Record<string, FnDef> = {
  AMORDEGRC: atp('x e e x x x j?', (cost: number, purchased: number, first: number, salvage: number, period: number, rate: number, basis = 0) => {
    if (!validBasis(basis) || rate <= 0 || cost < 0 || salvage < 0 || salvage > cost || period < 0 || purchased > first) return E.NUM
    const life = 1 / rate
    const coef = life < 3 ? 1 : life < 5 ? 1.5 : life <= 6 ? 2 : 2.5
    const r = rate * coef
    let dep = round(amorFirst(cost, purchased, first, r, basis), 0)
    let book = cost - dep
    let rest = book - salvage
    const n = Math.trunc(period)
    for (let k = 0; k < n; k++) {
      dep = round(r * book, 0)
      rest -= dep
      if (rest < 0) return n - k <= 1 ? round(book * 0.5, 0) : 0
      book -= dep
    }
    return dep
  }),
  AMORLINC: atp('x e e x x x j?', (cost: number, purchased: number, first: number, salvage: number, period: number, rate: number, basis = 0) => {
    if (!validBasis(basis) || rate <= 0 || cost < 0 || salvage < 0 || salvage > cost || period < 0 || purchased > first) return E.NUM
    const full = cost * rate
    const first0 = amorFirst(cost, purchased, first, rate, basis)
    const n = Math.trunc(period)
    if (n === 0) return first0
    const fulls = Math.floor((cost - salvage - first0) / full)
    if (n <= fulls) return full
    if (n === fulls + 1) return cost - salvage - full * fulls - first0
    return 0
  }),
  // the rate rounded to three decimals; a partial first year of `month` months and its remainder after `life`
  DB: {
    args: 'n n n n i?',
    fn: (cost: number, salvage: number, life: number, period: number, month = 12) => {
      if (month < 1 || month > 12 || life <= 0 || period < 1 || period > life + (month < 12 ? 1 : 0)) return E.NUM
      if (cost === 0) return 0
      const rate = round(1 - Math.pow(salvage / cost, 1 / life), 3)
      let total = (cost * rate * month) / 12
      let dep = total
      for (let p = 2; p <= Math.trunc(period); p++) {
        dep = (cost - total) * rate
        if (p > life) dep = (dep * (12 - month)) / 12
        total += dep
      }
      return dep
    },
  },
  DDB: {
    args: 'n n n n n?',
    fn: (cost: number, salvage: number, life: number, period: number, factor = 2) =>
      cost < 0 || salvage < 0 || life <= 0 || period <= 0 || period > life || factor <= 0 ? E.NUM : ddb(cost, salvage, life, period, factor),
  },
  SLN: { args: 'n n n', fn: (cost: number, salvage: number, life: number) => (life === 0 ? E.DIV0 : (cost - salvage) / life) },
  SYD: {
    args: 'n n n n',
    fn: (cost: number, salvage: number, life: number, per: number) =>
      salvage < 0 || life <= 0 || per <= 0 || per > life ? E.NUM : ((cost - salvage) * (life - per + 1) * 2) / (life * (life + 1)),
  },
  // fractional start and end take the matching share of their periods
  VDB: {
    args: 'n n n n n n? b?',
    fn: (cost: number, salvage: number, life: number, start: number, end: number, factor = 2, noSwitch = false) => {
      if (cost < 0 || salvage < 0 || life <= 0 || start < 0 || end < start || end > life || factor <= 0) return E.NUM
      const deps = schedule(cost, salvage, life, factor, noSwitch, Math.ceil(end))
      let sum = 0
      for (let p = Math.floor(start) + 1; p <= Math.ceil(end); p++) sum += deps[p - 1] * (Math.min(end, p) - Math.max(start, p - 1))
      return sum
    },
  },
}
