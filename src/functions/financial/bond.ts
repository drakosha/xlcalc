/** Bonds with regular coupons (PRICE, YIELD, DURATION, MDURATION) and ACCRINT, after Excel's documented formulas. */
import type { FnDef } from '../../args.js'
import { ymd } from '../../dates.js'
import { E } from '../../values.js'
import { atp } from '../date/atp.js'
import { validBasis } from '../date/daycount.js'
import { badCoupon, cmp, couponDate, couponDays, coupons, daysBetween, periodDays, periodOf } from './coupon.js'

/** Price per 100 at yield `yld`: cash flows discounted from settlement, less accrued interest. */
function price(settlement: number, maturity: number, rate: number, yld: number, redemption: number, freq: number, basis: number): number {
  const c = coupons(settlement, maturity, freq)
  const { e, a, dsc } = couponDays(c, basis, freq)
  const coupon = (100 * rate) / freq
  if (c.n === 1) return (redemption + coupon) / (1 + ((dsc / e) * yld) / freq) - (coupon * a) / e
  const v = 1 + yld / freq
  let p = redemption / Math.pow(v, c.n - 1 + dsc / e)
  for (let k = 1; k <= c.n; k++) p += coupon / Math.pow(v, k - 1 + dsc / e)
  return p - (coupon * a) / e
}

/** Macaulay duration in years. */
function duration(settlement: number, maturity: number, rate: number, yld: number, freq: number, basis: number): number {
  const c = coupons(settlement, maturity, freq)
  const { e, dsc } = couponDays(c, basis, freq)
  const v = 1 + yld / freq
  let pv = 0
  let weighted = 0
  for (let k = 1; k <= c.n; k++) {
    const t = k - 1 + dsc / e
    const cf = (100 * rate) / freq + (k === c.n ? 100 : 0)
    pv += cf / Math.pow(v, t)
    weighted += (t * cf) / Math.pow(v, t)
  }
  return weighted / pv / freq
}

/** A root of a decreasing `f` near `guess`: Newton's method with a numeric slope, bisection as the fallback. */
export function solveDecreasing(f: (x: number) => number, guess: number, lo: number): number | null {
  let x = guess
  for (let k = 0; k < 100; k++) {
    const y = f(x)
    if (!Number.isFinite(y)) break
    const h = Math.max(1e-6, Math.abs(x) * 1e-6)
    const slope = (f(x + h) - y) / h
    if (!(slope < 0)) break
    const next = Math.max((x + lo) / 2, x - y / slope)
    if (Math.abs(next - x) < 1e-14 * Math.max(1, Math.abs(x))) return next
    x = next
  }
  let a = lo
  let b = Math.max(1, guess)
  while (f(b) > 0 && b < 1e10) b *= 2
  if (!(f(a) >= 0) || f(b) > 0) return null
  for (let k = 0; k < 200 && b - a > 1e-15 * Math.max(1, Math.abs(b)); k++) {
    const m = (a + b) / 2
    if (f(m) > 0) a = m
    else b = m
  }
  return (a + b) / 2
}

export const functions: Record<string, FnDef> = {
  // the accrual from issue (or from first_interest, calc_method FALSE) in quasi-coupon periods anchored at first_interest; whole periods count 1
  ACCRINT: atp('e e e x x j j? b?', (issue: number, first: number, settlement: number, rate: number, par: number, freq: number, basis = 0, calc = true) => {
    if (rate <= 0 || par <= 0 || (freq !== 1 && freq !== 2 && freq !== 4) || !validBasis(basis) || issue >= settlement) return E.NUM
    const anchor = ymd(first)
    const start = ymd(calc || settlement <= first ? issue : first)
    const end = ymd(settlement)
    let sum = 0
    for (let k = periodOf(anchor, start, freq); ; k++) {
      const q0 = couponDate(anchor, k, freq)
      if (cmp(q0, end) >= 0) break
      const q1 = couponDate(anchor, k + 1, freq)
      const lo = cmp(start, q0) > 0 ? start : q0
      const hi = cmp(end, q1) < 0 ? end : q1
      sum += lo === q0 && hi === q1 ? 1 : daysBetween(lo, hi, basis) / periodDays(q0, q1, basis, freq)
    }
    return (par * rate * sum) / freq
  }),
  DURATION: atp('e e x x j j?', (settlement: number, maturity: number, rate: number, yld: number, freq: number, basis = 0) =>
    badCoupon(settlement, maturity, freq, basis) || rate < 0 || yld < 0 ? E.NUM : duration(settlement, maturity, rate, yld, freq, basis),
  ),
  MDURATION: atp('e e x x j j?', (settlement: number, maturity: number, rate: number, yld: number, freq: number, basis = 0) =>
    badCoupon(settlement, maturity, freq, basis) || rate < 0 || yld < 0 ? E.NUM : duration(settlement, maturity, rate, yld, freq, basis) / (1 + yld / freq),
  ),
  PRICE: atp('e e x x x j j?', (settlement: number, maturity: number, rate: number, yld: number, redemption: number, freq: number, basis = 0) =>
    badCoupon(settlement, maturity, freq, basis) || rate < 0 || yld < 0 || redemption <= 0 ? E.NUM : price(settlement, maturity, rate, yld, redemption, freq, basis),
  ),
  // one coupon left: Excel's closed form; else the yield at which PRICE gives `pr`
  YIELD: atp('e e x x x j j?', (settlement: number, maturity: number, rate: number, pr: number, redemption: number, freq: number, basis = 0) => {
    if (badCoupon(settlement, maturity, freq, basis) || rate < 0 || pr <= 0 || redemption <= 0) return E.NUM
    const c = coupons(settlement, maturity, freq)
    if (c.n === 1) {
      const { e, a, dsc } = couponDays(c, basis, freq)
      const paid = pr / 100 + ((a / e) * rate) / freq
      return ((redemption / 100 + rate / freq - paid) / paid) * ((freq * e) / dsc)
    }
    return solveDecreasing((y) => price(settlement, maturity, rate, y, redemption, freq, basis) - pr, rate || 0.05, -freq + 1e-9) ?? E.NUM
  }),
}
