/** Bonds with an odd first or last period (ODDF*, ODDL*), after Excel's documented formulas over quasi-coupon periods. */
import type { FnDef } from '../../args.js'
import { ymd } from '../../dates.js'
import { E, type CErr } from '../../values.js'
import { atp } from '../date/atp.js'
import { validBasis } from '../date/daycount.js'
import { solveDecreasing } from './bond.js'
import { cmp, couponDate, daysBetween, periodDays, periodOf, type Ymd } from './coupon.js'

const badFreq = (freq: number, basis: number) => (freq !== 1 && freq !== 2 && freq !== 4) || !validBasis(basis)
const later = (a: Ymd, b: Ymd) => (cmp(a, b) > 0 ? a : b)
const earlier = (a: Ymd, b: Ymd) => (cmp(a, b) < 0 ? a : b)

/** ODDFPRICE at yield `yld`: quasi-coupon periods counted back from the first coupon. */
function oddFirstPrice(settlement: number, maturity: number, issue: number, first: number, rate: number, yld: number, redemption: number, freq: number, basis: number): number {
  const f1 = ymd(first)
  const iss = ymd(issue)
  const settle = ymd(settlement)
  const q = (k: number) => couponDate(f1, k, freq)
  const coupon = (100 * rate) / freq
  const n = periodOf(f1, ymd(maturity), freq)
  const ki = periodOf(f1, iss, freq)
  const ks = periodOf(f1, settle, freq)
  let odd = 0
  let accrued = 0
  for (let k = ki; k < 0; k++) {
    const nl = periodDays(q(k), q(k + 1), basis, freq)
    const from = later(iss, q(k))
    odd += daysBetween(from, q(k + 1), basis) / nl
    if (k <= ks) accrued += daysBetween(from, earlier(settle, q(k + 1)), basis) / nl
  }
  const t = -(ks + 1) + daysBetween(settle, q(ks + 1), basis) / periodDays(q(ks), q(ks + 1), basis, freq)
  const v = 1 + yld / freq
  let p = (coupon * odd) / Math.pow(v, t) + redemption / Math.pow(v, t + n)
  for (let j = 1; j <= n; j++) p += coupon / Math.pow(v, t + j)
  return p - coupon * accrued
}

/** ODDLPRICE and ODDLYIELD: sums over quasi-coupon periods counted on from the last coupon. */
function oddLast(settlement: number, maturity: number, last: number, freq: number, basis: number) {
  const l = ymd(last)
  const settle = ymd(settlement)
  const mat = ymd(maturity)
  let dc = 0
  let a = 0
  let dsc = 0
  for (let k = 0; cmp(couponDate(l, k, freq), mat) < 0; k++) {
    const q0 = couponDate(l, k, freq)
    const q1 = couponDate(l, k + 1, freq)
    const nl = periodDays(q0, q1, basis, freq)
    const end = earlier(q1, mat)
    dc += daysBetween(q0, end, basis) / nl
    if (cmp(settle, q0) > 0) a += daysBetween(q0, earlier(settle, end), basis) / nl
    if (cmp(settle, end) < 0) dsc += daysBetween(later(settle, q0), end, basis) / nl
  }
  return { dc, a, dsc }
}

const oddFirstBad = (settlement: number, maturity: number, issue: number, first: number, freq: number, basis: number) =>
  badFreq(freq, basis) || !(issue < settlement && settlement < first && first < maturity)

const oddLastBad = (settlement: number, maturity: number, last: number, freq: number, basis: number) =>
  badFreq(freq, basis) || !(last < settlement && settlement < maturity)

export const functions: Record<string, FnDef> = {
  ODDFPRICE: atp('e e e e x x x j j?', (settlement: number, maturity: number, issue: number, first: number, rate: number, yld: number, redemption: number, freq: number, basis = 0) =>
    oddFirstBad(settlement, maturity, issue, first, freq, basis) || rate < 0 || yld < 0 || redemption <= 0
      ? E.NUM
      : oddFirstPrice(settlement, maturity, issue, first, rate, yld, redemption, freq, basis),
  ),
  ODDFYIELD: atp('e e e e x x x j j?', (settlement: number, maturity: number, issue: number, first: number, rate: number, pr: number, redemption: number, freq: number, basis = 0): number | CErr => {
    if (oddFirstBad(settlement, maturity, issue, first, freq, basis) || rate < 0 || pr <= 0 || redemption <= 0) return E.NUM
    const f = (y: number) => oddFirstPrice(settlement, maturity, issue, first, rate, y, redemption, freq, basis) - pr
    return solveDecreasing(f, rate || 0.05, -freq + 1e-9) ?? E.NUM
  }),
  ODDLPRICE: atp('e e e x x x j j?', (settlement: number, maturity: number, last: number, rate: number, yld: number, redemption: number, freq: number, basis = 0) => {
    if (oddLastBad(settlement, maturity, last, freq, basis) || rate < 0 || yld < 0 || redemption <= 0) return E.NUM
    const { dc, a, dsc } = oddLast(settlement, maturity, last, freq, basis)
    const coupon = (100 * rate) / freq
    return (redemption + dc * coupon) / (1 + (dsc * yld) / freq) - a * coupon
  }),
  ODDLYIELD: atp('e e e x x x j j?', (settlement: number, maturity: number, last: number, rate: number, pr: number, redemption: number, freq: number, basis = 0) => {
    if (oddLastBad(settlement, maturity, last, freq, basis) || rate < 0 || pr <= 0 || redemption <= 0) return E.NUM
    const { dc, a, dsc } = oddLast(settlement, maturity, last, freq, basis)
    const coupon = (100 * rate) / freq
    const paid = pr + a * coupon
    return ((redemption + dc * coupon - paid) / paid) * (freq / dsc)
  }),
}
