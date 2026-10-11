/** Day-count conventions (DAYS360, YEARFRAC), shared with the financial functions. */
import type { FnDef } from '../../args.js'
import { isLeap, lastDayExcel, serialOf, utcDays, ymd } from '../../dates.js'
import { atp, need, type AtpOrder } from './atp.js'

export type Ymd = { y: number; m: number; d: number }

/** Days from `a` to `b` on 30/360, their days taken as `d1` and `d2`: each convention moves them by its own rules. */
export const d360 = (a: Ymd, b: Ymd, d1: number, d2: number) => (b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1)

/** DAYS360: US moves a start on the 31st or on the last of February (Excel's calendar) to the 30th. */
export function days360(start: number, end: number, european: boolean): number {
  const a = ymd(start)
  const b = ymd(end)
  let d1 = a.d
  let d2 = b.d
  if (european) {
    d1 = Math.min(d1, 30)
    d2 = Math.min(d2, 30)
  } else {
    if (d1 === 31 || (a.m === 2 && d1 === lastDayExcel(a.y, 2))) d1 = 30
    if (d2 === 31 && d1 >= 30) d2 = 30
  }
  return d360(a, b, d1, d2)
}

/** By the real calendar: the ToolPak takes 1900 for a common year. */
const lastOfFeb = (x: Ymd) => x.m === 2 && x.d === (isLeap(x.y) ? 29 : 28)

/** The ToolPak's 30/360 (YEARFRAC basis 0) between ordered dates: NASD with its February rules. */
export function days30(a: Ymd, b: Ymd): number {
  let d1 = a.d
  let d2 = b.d
  if (d1 === 31 && d2 === 31) [d1, d2] = [30, 30]
  else if (d1 === 31) d1 = 30
  else if (d1 === 30 && d2 === 31) d2 = 30
  else if (lastOfFeb(a) && lastOfFeb(b)) [d1, d2] = [30, 30]
  else if (lastOfFeb(a)) d1 = 30
  return d360(a, b, d1, d2)
}

/** The actual/actual year between ordered dates: 365 or 366 within a year, else the span's average by the real calendar. */
export function yearLength(start: number, end: number): number {
  const a = ymd(start)
  const b = ymd(end)
  if (a.y === b.y || (b.y === a.y + 1 && (a.m > b.m || (a.m === b.m && a.d >= b.d)))) {
    const mar1a = serialOf(a.y, 3, 1)
    const mar1b = serialOf(b.y, 3, 1)
    const leap =
      a.y === b.y
        ? isLeap(a.y)
        : (isLeap(a.y) && start < mar1a && end >= mar1a) || (isLeap(b.y) && end >= mar1b && start < mar1b)
    return leap || (b.m === 2 && b.d === 29) ? 366 : 365
  }
  // by the real calendar, 1900 a common year (`utcDays`, no phantom day)
  return (utcDays(b.y + 1, 1, 1) - utcDays(a.y, 1, 1)) / (b.y - a.y + 1)
}

/** Years between two dates (either order) by basis: 0 US 30/360, 1 actual/actual, 2 actual/360, 3 actual/365, 4 European 30/360. */
export function yearFrac(start: number, end: number, basis: number): number {
  if (start > end) [start, end] = [end, start]
  if (basis === 0) return days30(ymd(start), ymd(end)) / 360
  if (basis === 1) return (end - start) / yearLength(start, end)
  if (basis === 2) return (end - start) / 360
  if (basis === 3) return (end - start) / 365
  return days360(start, end, true) / 360
}

export const validBasis = (b: number) => b >= 0 && b <= 4

export const validFreq = (f: number) => f === 1 || f === 2 || f === 4

/** Excel reads the basis first in YEARFRAC and most securities: an invalid one is #NUM! whatever the other arguments. */
export const basisFirst = (i: number): AtpOrder => ({ first: [i], check: { [i]: need(validBasis) } })

/** The frequency of a security is checked as it is read, before the basis (DURATION, ACCRINT, ODDFPRICE). */
export const freqAt = (i: number): AtpOrder => ({ check: { [i]: need(validFreq) } })

export const functions: Record<string, FnDef> = {
  DAYS360: { args: 'd d b?', fn: (start: number, end: number, european = false) => days360(start, end, european) },
  YEARFRAC: atp('e e j?', (start: number, end: number, basis = 0) => yearFrac(start, end, basis), basisFirst(2)),
}
