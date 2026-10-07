/** Coupon schedules and the COUP* functions, on Excel's serial numbers. */
import type { FnDef } from '../../args.js'
import { addMonths, daysInMonth, serialOf, ymd } from '../../dates.js'
import { E, type CErr } from '../../values.js'
import { atp, need } from '../date/atp.js'
import { days30, validBasis, validFreq, type Ymd } from '../date/daycount.js'

export type { Ymd }

export const cmp = (a: Ymd, b: Ymd) => a.y - b.y || a.m - b.m || a.d - b.d

/** Excel's serial number of a date, 1900-02-29 included. */
export const serialOfYmd = (x: Ymd) => (x.y === 1900 && x.m === 2 && x.d === 29 ? 60 : serialOf(x.y, x.m, x.d))

/** Days from `a` to `b` by basis: 0 the ToolPak's 30/360, 4 European 30/360, else actual. */
export function daysBetween(a: Ymd, b: Ymd, basis: number): number {
  if (basis === 0) return cmp(a, b) <= 0 ? days30(a, b) : -days30(b, a)
  if (basis === 4) return (b.y - a.y) * 360 + (b.m - a.m) * 30 + Math.min(b.d, 30) - Math.min(a.d, 30)
  return serialOfYmd(b) - serialOfYmd(a)
}

/** Coupon dates `k` periods from `anchor`; an anchor on a month's last day keeps months' last days (unless `eom` is false). */
export function couponDate(anchor: Ymd, k: number, freq: number, eom = true): Ymd {
  const t = addMonths(anchor.y, anchor.m, (k * 12) / freq)
  const last = daysInMonth(t.y, t.m)
  return { y: t.y, m: t.m, d: eom && anchor.d === daysInMonth(anchor.y, anchor.m) ? last : Math.min(anchor.d, last) }
}

/** The coupon period around `x` on the schedule of `anchor`: k with date(k) ≤ x < date(k+1). */
export function periodOf(anchor: Ymd, x: Ymd, freq: number): number {
  let k = Math.floor((((x.y - anchor.y) * 12 + (x.m - anchor.m)) * freq) / 12)
  while (cmp(couponDate(anchor, k, freq), x) > 0) k--
  while (cmp(couponDate(anchor, k + 1, freq), x) <= 0) k++
  return k
}

/** The normal coupon period in days on the 30/360 and actual/360, actual/365 bases. */
export const normalDays = (basis: number, freq: number) => (basis === 3 ? 365 : 360) / freq

export interface Coupons {
  settlement: number
  settle: Ymd
  /** Before serial 0 the previous coupon date is day 0 (1900-01-00), as in Excel (`early`). */
  pcd: Ymd
  early?: boolean
  ncd: Ymd
  /** Coupons from settlement to maturity; 0 when the previous coupon date is before 1900. */
  n: number
}

export function coupons(settlement: number, maturity: number, freq: number): Coupons {
  const settle = ymd(settlement)
  const mat = ymd(maturity)
  const k = periodOf(mat, settle, freq)
  const pcd = couponDate(mat, k, freq)
  const ncd = couponDate(mat, k + 1, freq)
  return serialOfYmd(pcd) < 0 ? { settlement, settle, pcd: ymd(0), early: true, ncd, n: 0 } : { settlement, settle, pcd, ncd, n: -k }
}

/** Argument checks shared by the coupon and bond functions. */
export const badCoupon = (settlement: number, maturity: number, freq: number, basis: number) =>
  settlement >= maturity || !validFreq(freq) || !validBasis(basis)

/** Days from a coupon date to settlement (A) by basis; actual days are serial differences. */
const daysFrom = (c: Coupons, basis: number) => (basis === 0 || basis === 4 ? daysBetween(c.pcd, c.settle, basis) : c.settlement - serialOfYmd(c.pcd))

/** 30/360 with no month-end rule: Excel's COUPDAYSNC on basis 0 when the previous coupon date is before 1900. */
const plain360 = (a: Ymd, b: Ymd) => (b.y - a.y) * 360 + (b.m - a.m) * 30 + b.d - a.d

/** A and DSC as COUPDAYBS and COUPDAYSNC answer them; DSC on basis 0 is the normal period less A. */
export function couponDays(c: Coupons, basis: number, freq: number) {
  const a = daysFrom(c, basis)
  let dsc: number
  if (basis === 0) dsc = c.early ? plain360(c.settle, c.ncd) : normalDays(0, freq) - a
  else dsc = basis === 4 ? daysBetween(c.settle, c.ncd, 4) : serialOfYmd(c.ncd) - c.settlement
  return { a, dsc }
}

/** E, A and DSC as PRICE, YIELD and DURATION take them: E on basis 1 is the actual period, DSC = E − A. */
export function bondDays(c: Coupons, basis: number, freq: number) {
  const e = basis === 1 ? serialOfYmd(c.ncd) - serialOfYmd(c.pcd) : normalDays(basis, freq)
  const a = daysFrom(c, basis)
  return { e, a, dsc: e - a }
}

/** The COUP* functions check each argument as they read it: the dates' order, then the frequency, then the basis. */
function coup(answer: (c: Coupons, basis: number, freq: number) => number | CErr, date?: FnDef['date']): FnDef {
  return atp('e e j j?', (settlement: number, maturity: number, freq: number, basis = 0) => answer(coupons(settlement, maturity, freq), basis, freq), {
    check: { 1: need((maturity, a) => a[0] < maturity), 2: need(validFreq), 3: need(validBasis) },
    date,
  })
}

export const functions: Record<string, FnDef> = {
  COUPDAYBS: coup((c, basis, freq) => couponDays(c, basis, freq).a),
  COUPDAYSNC: coup((c, basis, freq) => couponDays(c, basis, freq).dsc),
  COUPNCD: coup((c) => serialOfYmd(c.ncd), 'calendar'),
  COUPNUM: coup((c) => c.n || E.NUM),
  COUPPCD: coup((c) => serialOfYmd(c.pcd), 'calendar'),
}
