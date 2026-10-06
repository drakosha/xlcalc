/** Bonds with an odd first or last period (ODDF*, ODDL*), after Excel's documented formulas over quasi-coupon periods. */
import type { FnDef } from '../../args.js'
import { ymd } from '../../dates.js'
import { E, type CErr } from '../../values.js'
import { atp, need } from '../date/atp.js'
import { validBasis, validFreq } from '../date/daycount.js'
import { solveDecreasing } from './bond.js'
import { cmp, couponDate, daysBetween, normalDays, periodOf, type Ymd } from './coupon.js'

const badFreq = (freq: number, basis: number) => !validFreq(freq) || !validBasis(basis)

/** The frequency is checked as it is read, before the basis. */
const freqAt = (i: number) => ({ check: { [i]: need(validFreq) } })
const later = (a: Ymd, b: Ymd) => (cmp(a, b) > 0 ? a : b)

/** A quasi-coupon period's normal length (`actual`: bases 1–3 count actual days, as ODDL* do in Excel). */
const normal = (q0: Ymd, q1: Ymd, basis: number, freq: number, actual = basis === 1) => (actual ? daysBetween(q0, q1, 1) : normalDays(basis, freq))
const earlier = (a: Ymd, b: Ymd) => (cmp(a, b) < 0 ? a : b)

/** ODDFPRICE at yield `yld`: quasi-coupon periods counted back from the first coupon. */
function oddFirstPrice(settlement: number, maturity: number, issue: number, first: number, rate: number, yld: number, redemption: number, freq: number, basis: number): number {
  const f1 = ymd(first)
  const iss = ymd(issue)
  const settle = ymd(settlement)
  // memoized: whole periods are told by identity below
  const dates = new Map<number, Ymd>()
  const q = (k: number) => {
    let d = dates.get(k)
    if (!d) dates.set(k, (d = couponDate(f1, k, freq)))
    return d
  }
  const coupon = (100 * rate) / freq
  const n = periodOf(f1, ymd(maturity), freq)
  const ki = periodOf(f1, iss, freq)
  const ks = periodOf(f1, settle, freq)
  let odd = 0
  let accrued = 0
  for (let k = ki; k < 0; k++) {
    const nl = normal(q(k), q(k + 1), basis, freq)
    const from = later(iss, q(k))
    // a whole quasi-coupon period adds one coupon to the odd one; accrued interest still counts its days
    odd += from === q(k) ? 1 : daysBetween(from, q(k + 1), basis) / nl
    if (k <= ks) accrued += daysBetween(from, earlier(settle, q(k + 1)), basis) / nl
  }
  const t = -(ks + 1) + daysBetween(settle, q(ks + 1), basis) / normal(q(ks), q(ks + 1), basis, freq)
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
  // quasi-coupon dates keep the last coupon's day, month ends or not
  for (let k = 0; cmp(couponDate(l, k, freq, false), mat) < 0; k++) {
    const q0 = couponDate(l, k, freq, false)
    const q1 = couponDate(l, k + 1, freq, false)
    const nl = normal(q0, q1, basis, freq, basis >= 1 && basis <= 3)
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
  freqAt(7)),
  ODDFYIELD: atp('e e e e x x x j j?', (settlement: number, maturity: number, issue: number, first: number, rate: number, pr: number, redemption: number, freq: number, basis = 0): number | CErr => {
    if (oddFirstBad(settlement, maturity, issue, first, freq, basis) || rate < 0 || pr <= 0 || redemption <= 0) return E.NUM
    const f = (y: number) => oddFirstPrice(settlement, maturity, issue, first, rate, y, redemption, freq, basis) - pr
    return solveDecreasing(f, rate || 0.05, -freq + 1e-9) ?? E.NUM
  }, freqAt(7)),
  ODDLPRICE: atp('e e e x x x j j?', (settlement: number, maturity: number, last: number, rate: number, yld: number, redemption: number, freq: number, basis = 0) => {
    if (oddLastBad(settlement, maturity, last, freq, basis) || rate < 0 || yld < 0 || redemption <= 0) return E.NUM
    const { dc, a, dsc } = oddLast(settlement, maturity, last, freq, basis)
    const coupon = (100 * rate) / freq
    return (redemption + dc * coupon) / (1 + (dsc * yld) / freq) - a * coupon
  }, freqAt(6)),
  ODDLYIELD: atp('e e e x x x j j?', (settlement: number, maturity: number, last: number, rate: number, pr: number, redemption: number, freq: number, basis = 0) => {
    if (oddLastBad(settlement, maturity, last, freq, basis) || rate < 0 || pr <= 0 || redemption <= 0) return E.NUM
    const { dc, a, dsc } = oddLast(settlement, maturity, last, freq, basis)
    const coupon = (100 * rate) / freq
    const paid = pr + a * coupon
    return ((redemption + dc * coupon - paid) / paid) * (freq / dsc)
  }, freqAt(6)),
}
