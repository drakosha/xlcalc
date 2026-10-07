/** Bonds with regular coupons (PRICE, YIELD, DURATION, MDURATION) and ACCRINT, after Excel's documented formulas. */
import type { FnDef } from '../../args.js'
import { ymd } from '../../dates.js'
import { E } from '../../values.js'
import { atp, need } from '../date/atp.js'
import { basisFirst, validBasis, validFreq } from '../date/daycount.js'
import { addMonths, daysInMonth } from '../../dates.js'
import { badCoupon, bondDays, cmp, couponDate, coupons, daysBetween, normalDays, periodOf, serialOfYmd, type Ymd } from './coupon.js'

/** The date `months` earlier, the day clamped to the month. */
function monthsBack(x: Ymd, months: number): Ymd {
  const t = addMonths(x.y, x.m, -months)
  return { y: t.y, m: t.m, d: Math.min(x.d, daysInMonth(t.y, t.m)) }
}

/** Price per 100 at yield `yld`: cash flows discounted from settlement, less accrued interest. */
function price(settlement: number, maturity: number, rate: number, yld: number, redemption: number, freq: number, basis: number): number {
  const c = coupons(settlement, maturity, freq)
  const { e, a, dsc } = bondDays(c, basis, freq)
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
  const { e, dsc } = bondDays(c, basis, freq)
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

/** DURATION and MDURATION check the frequency as they read it; PRICE and YIELD read the basis first. */
const FREQ_READ = { check: { 4: need(validFreq) } }
const PRICE_READ = { first: [6], check: { 6: need(validBasis), 5: need(validFreq) } }

export const functions: Record<string, FnDef> = {
  // quasi-coupon periods anchored at first_interest, whole ones counting 1; calc_method FALSE counts the days in one piece
  ACCRINT: atp('e e e x x j j? b?', (issue: number, first: number, settlement: number, rate: number, par = 1000, freq: number, basis = 0, calc = true) => {
    if (rate <= 0 || par <= 0 || !validBasis(basis) || issue >= settlement) return E.NUM
    const anchor = ymd(first)
    const start = ymd(issue)
    const end = ymd(settlement)
    const normal = normalDays(basis, freq)
    // the odd first period's length counts its days on bases 1 and 4; later ones on basis 1 only, back from their end
    const nl = (k: number, q0: Ymd, q1: Ymd) => {
      if (basis === 1) return k < 0 ? daysBetween(q0, q1, 1) : serialOfYmd(q1) - serialOfYmd(monthsBack(q1, 12 / freq))
      return basis === 4 && k < 0 && cmp(end, q1) >= 0 ? daysBetween(q0, q1, 4) : normal
    }
    if (!calc && settlement > first) return (par * rate * daysBetween(start, end, basis)) / freq / (basis === 1 ? nl(-1, couponDate(anchor, -1, freq), anchor) : normal)
    let sum = 0
    for (let k = periodOf(anchor, start, freq); ; k++) {
      const q0 = couponDate(anchor, k, freq)
      if (cmp(q0, end) >= 0) break
      const q1 = couponDate(anchor, k + 1, freq)
      const lo = cmp(start, q0) > 0 ? start : q0
      const hi = cmp(end, q1) < 0 ? end : q1
      sum += lo === q0 && hi === q1 ? 1 : daysBetween(lo, hi, basis) / nl(k, q0, q1)
    }
    return (par * rate * sum) / freq
  }, { check: { 5: need(validFreq) } }),
  DURATION: atp('e e x x j j?', (settlement: number, maturity: number, rate: number, yld: number, freq: number, basis = 0) =>
    badCoupon(settlement, maturity, freq, basis) || rate < 0 || yld < 0 ? E.NUM : duration(settlement, maturity, rate, yld, freq, basis),
  FREQ_READ),
  MDURATION: atp('e e x x j j?', (settlement: number, maturity: number, rate: number, yld: number, freq: number, basis = 0) =>
    badCoupon(settlement, maturity, freq, basis) || rate < 0 || yld < 0 ? E.NUM : duration(settlement, maturity, rate, yld, freq, basis) / (1 + yld / freq),
  FREQ_READ),
  PRICE: atp('e e x x x j j?', (settlement: number, maturity: number, rate: number, yld: number, redemption: number, freq: number, basis = 0) =>
    badCoupon(settlement, maturity, freq, basis) || rate < 0 || yld < 0 || redemption <= 0 ? E.NUM : price(settlement, maturity, rate, yld, redemption, freq, basis),
  PRICE_READ),
  // one coupon left: Excel's closed form; else the yield at which PRICE gives `pr`
  YIELD: atp('e e x x x j j?', (settlement: number, maturity: number, rate: number, pr: number, redemption: number, freq: number, basis = 0) => {
    if (badCoupon(settlement, maturity, freq, basis) || rate < 0 || pr <= 0 || redemption <= 0) return E.NUM
    const c = coupons(settlement, maturity, freq)
    if (c.n === 1) {
      const { e, a, dsc } = bondDays(c, basis, freq)
      const paid = pr / 100 + ((a / e) * rate) / freq
      return ((redemption / 100 + rate / freq - paid) / paid) * ((freq * e) / dsc)
    }
    return solveDecreasing((y) => price(settlement, maturity, rate, y, redemption, freq, basis) - pr, rate || 0.05, -freq + 1e-9) ?? E.NUM
  }, PRICE_READ),
}
