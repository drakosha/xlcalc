/** Coupon schedules and the COUP* functions. Coupon dates live in the real calendar (1900 is a common year). */
import type { FnDef } from '../../args.js'
import { addMonths, daysInMonth, serialOf, ymd } from '../../dates.js'
import { E, type CErr } from '../../values.js'
import { atp } from '../date/atp.js'
import { dayNum, days30, validBasis, type Ymd } from '../date/daycount.js'

export type { Ymd }

export const cmp = (a: Ymd, b: Ymd) => a.y - b.y || a.m - b.m || a.d - b.d

/** Days from `a` to `b` by basis: 0 the ToolPak's 30/360, 4 European 30/360, else actual. */
export function daysBetween(a: Ymd, b: Ymd, basis: number): number {
  if (basis === 0) return cmp(a, b) <= 0 ? days30(a, b) : -days30(b, a)
  if (basis === 4) return (b.y - a.y) * 360 + (b.m - a.m) * 30 + Math.min(b.d, 30) - Math.min(a.d, 30)
  return dayNum(b) - dayNum(a)
}

/** The normal length of the coupon period [a, b] in days. */
export const periodDays = (a: Ymd, b: Ymd, basis: number, freq: number) => (basis === 1 ? dayNum(b) - dayNum(a) : (basis === 3 ? 365 : 360) / freq)

/** Coupon dates `k` periods from `anchor`; an anchor on a month's last day keeps months' last days. */
export function couponDate(anchor: Ymd, k: number, freq: number): Ymd {
  const t = addMonths(anchor.y, anchor.m, (k * 12) / freq)
  const last = daysInMonth(t.y, t.m)
  return { y: t.y, m: t.m, d: anchor.d === daysInMonth(anchor.y, anchor.m) ? last : Math.min(anchor.d, last) }
}

/** The coupon period around `x` on the schedule of `anchor`: k with date(k) ≤ x < date(k+1). */
export function periodOf(anchor: Ymd, x: Ymd, freq: number): number {
  let k = Math.floor((((x.y - anchor.y) * 12 + (x.m - anchor.m)) * freq) / 12)
  while (cmp(couponDate(anchor, k, freq), x) > 0) k--
  while (cmp(couponDate(anchor, k + 1, freq), x) <= 0) k++
  return k
}

export interface Coupons {
  settle: Ymd
  pcd: Ymd
  ncd: Ymd
  /** Coupons from settlement to maturity. */
  n: number
}

export function coupons(settlement: number, maturity: number, freq: number): Coupons {
  const settle = ymd(settlement)
  const mat = ymd(maturity)
  const k = periodOf(mat, settle, freq)
  return { settle, pcd: couponDate(mat, k, freq), ncd: couponDate(mat, k + 1, freq), n: -k }
}

/** Argument checks shared by the coupon and bond functions. */
export const badCoupon = (settlement: number, maturity: number, freq: number, basis: number) =>
  settlement >= maturity || (freq !== 1 && freq !== 2 && freq !== 4) || !validBasis(basis)

/** COUPDAYS, COUPDAYBS, COUPDAYSNC: the period length E, the days before settlement A and after it DSC. */
export function couponDays(c: Coupons, basis: number, freq: number) {
  return { e: periodDays(c.pcd, c.ncd, basis, freq), a: daysBetween(c.pcd, c.settle, basis), dsc: daysBetween(c.settle, c.ncd, basis) }
}

function coup(answer: (c: Coupons, basis: number, freq: number) => number | CErr): FnDef {
  return atp('e e j j?', (settlement: number, maturity: number, freq: number, basis = 0) =>
    badCoupon(settlement, maturity, freq, basis) ? E.NUM : answer(coupons(settlement, maturity, freq), basis, freq),
  )
}

const serialOfYmd = (x: Ymd) => serialOf(x.y, x.m, x.d)

export const functions: Record<string, FnDef> = {
  COUPDAYBS: coup((c, basis, freq) => couponDays(c, basis, freq).a),
  COUPDAYS: coup((c, basis, freq) => couponDays(c, basis, freq).e),
  COUPDAYSNC: coup((c, basis, freq) => couponDays(c, basis, freq).dsc),
  COUPNCD: coup((c) => serialOfYmd(c.ncd)),
  COUPNUM: coup((c) => c.n),
  COUPPCD: coup((c) => serialOfYmd(c.pcd)),
}
