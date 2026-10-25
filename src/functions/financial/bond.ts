/** Bonds with regular coupons (PRICE, YIELD, DURATION, MDURATION) and ACCRINT, after Excel's documented formulas. */
import type { FnDef } from '../../args.js'
import { addMonths, daysInMonth, ymd } from '../../dates.js'
import { E } from '../../values.js'
import { atp, need } from '../date/atp.js'
import { freqAt, euro360, us360, validBasis, validFreq } from '../date/daycount.js'
import { badCoupon, bondDays, cmp, coupons, normalDays, serialOfYmd, type Ymd } from './coupon.js'

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

/** A coupon date `n` months from `x` as ACCRINT steps its schedule: the day kept within the month, an end-of-month bond's on the last. */
function shift(x: Ymd, n: number, eom: boolean): Ymd {
  const t = addMonths(x.y, x.m, n)
  const last = daysInMonth(t.y, t.m)
  return { y: t.y, m: t.m, d: eom ? last : Math.min(x.d, last) }
}

/**
 * ACCRINT's coupon periods, as Excel counts them (the algorithm of ExcelFinancialFunctions): dates stepped from
 * first_interest, each from the one before; the period before the settlement counts its days over the length of the
 * period ending at first_interest, each whole period back to the issue 1 (0 with calc_method FALSE, and the settlement's
 * period then starts at that regular one), the issue's own period its days over that period's length.
 */
function accrint(issue: number, first: number, settlement: number, freq: number, basis: number, calc: boolean): number {
  const months = 12 / freq
  const fi = ymd(first)
  const is = ymd(issue)
  const st = ymd(settlement)
  const eom = fi.d === daysInMonth(fi.y, fi.m)
  const days = (a: Ymd, b: Ymd) => (basis === 0 ? us360(a, b, false) : basis === 4 ? euro360(a, b) : serialOfYmd(b) - serialOfYmd(a))
  const regular = shift(fi, -months, eom)
  let pcd = regular
  if (calc && cmp(st, fi) > 0) for (let q = fi; cmp(q, st) < 0; q = shift(q, months, eom)) pcd = q
  let sum = days(cmp(is, pcd) > 0 ? is : pcd, st) / (basis === 1 ? serialOfYmd(fi) - serialOfYmd(regular) : normalDays(basis, freq))
  for (let end = pcd; cmp(end, is) > 0; ) {
    const start = shift(end, -months, eom)
    if (cmp(is, start) <= 0) sum += calc ? 1 : 0
    // the issue's period beginning before 1900 leaves no periods to count (`ACCRINT(1,…)` is 0)
    else if (serialOfYmd(start) < 0) return 0
    else sum += days(is, end) / (basis === 0 ? us360(start, end, true) : basis === 2 || basis === 3 ? normalDays(basis, freq) : days(start, end))
    end = start
  }
  return sum
}

/** DURATION and MDURATION check the frequency as they read it; PRICE and YIELD read the basis first. */
const FREQ_READ = freqAt(4)
const PRICE_READ = { first: [6], check: { 6: need(validBasis), 5: need(validFreq) } }

export const functions: Record<string, FnDef> = {
  ACCRINT: atp('e e e x x j j? b?', (issue: number, first: number, settlement: number, rate: number, par = 1000, freq: number, basis = 0, calc = true) =>
    rate <= 0 || par <= 0 || !validBasis(basis) || issue >= settlement || issue < 1 || first < 1 ? E.NUM : accrint(issue, first, settlement, freq, basis, calc) * ((par * rate) / freq),
  freqAt(5)),
  DURATION: atp('e e x x j j?', (settlement: number, maturity: number, rate: number, yld: number, freq: number, basis = 0) =>
    badCoupon(settlement, maturity, freq, basis) || rate < 0 || yld < 0 ? E.NUM : duration(settlement, maturity, rate, yld, freq, basis),
  FREQ_READ),
  MDURATION: atp('e e x x j j?', (settlement: number, maturity: number, rate: number, yld: number, freq: number, basis = 0) =>
    badCoupon(settlement, maturity, freq, basis) || rate < 0 || yld < 0 ? E.NUM : duration(settlement, maturity, rate, yld, freq, basis) / (1 + yld / freq),
  FREQ_READ),
  // a previous coupon date before 1900 leaves no coupons, as COUPNUM finds none
  PRICE: atp('e e x x x j j?', (settlement: number, maturity: number, rate: number, yld: number, redemption: number, freq: number, basis = 0) =>
    badCoupon(settlement, maturity, freq, basis) || rate < 0 || yld < 0 || redemption <= 0 || coupons(settlement, maturity, freq).early ? E.NUM : price(settlement, maturity, rate, yld, redemption, freq, basis),
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
