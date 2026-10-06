/** Workdays: NETWORKDAYS and WORKDAY with weekend codes and holidays. */
import type { FnDef } from '../../args.js'
import { MAX_SERIAL } from '../../dates.js'
import { E, isErr, type CErr, type V } from '../../values.js'
import { atp } from './atp.js'
import { weekday } from './calendar.js'

/** Weekend mask (index 0 Sunday) by code 1–7, 11–17 or seven 0/1 from Monday; an error there is #NUM! in Excel. */
function weekendOf(w: V | undefined): boolean[] | CErr {
  const mask = [false, false, false, false, false, false, false]
  if (w === undefined || w === null) w = 1
  if (isErr(w)) return E.NUM
  if (typeof w === 'boolean') return E.VALUE
  if (typeof w === 'string') {
    if (!/^[01]{7}$/.test(w) || w === '1111111') return E.VALUE
    for (let i = 0; i < 7; i++) mask[(i + 1) % 7] = w[i] === '1'
    return mask
  }
  const k = Math.trunc(w)
  if (k >= 1 && k <= 7) {
    mask[(k + 5) % 7] = true
    mask[(k + 6) % 7] = true
  } else if (k >= 11 && k <= 17) mask[k - 11] = true
  else return E.NUM
  return mask
}

const SAT_SUN = weekendOf(1) as boolean[]

/** Inclusive count, negative when `end` is earlier. */
function networkDays(start: number, end: number, weekend: boolean[], holidays: number[] = []): number {
  const sign = start <= end ? 1 : -1
  const [a, b] = sign > 0 ? [start, end] : [end, start]
  const weeks = Math.floor((b - a + 1) / 7)
  let count = weeks * weekend.filter((x) => !x).length
  for (let s = a + weeks * 7; s <= b; s++) if (!weekend[weekday(s)]) count++
  for (const h of new Set(holidays)) if (h >= a && h <= b && !weekend[weekday(h)]) count--
  return sign * count
}

/** Whole weeks first, then day by day; holidays passed on workdays are made up past the date reached. */
function workday(start: number, days: number, weekend: boolean[], holidays: number[] = []): number | CErr {
  if (days === 0) return start
  if (Math.abs(days) > MAX_SERIAL) return E.NUM
  const hol = new Set(holidays)
  const step = days > 0 ? 1 : -1
  const perWeek = weekend.filter((x) => !x).length
  const weeks = Math.floor((Math.abs(days) - 1) / perWeek)
  let left = Math.abs(days) - weeks * perWeek
  let d = start + step * weeks * 7
  while (left > 0) {
    d += step
    if (!weekend[weekday(d)]) left--
  }
  let extra = 0
  for (const h of hol) if ((step > 0 ? h > start && h <= d : h < start && h >= d) && !weekend[weekday(h)]) extra++
  while (extra > 0) {
    d += step
    if (d < 0 || d > MAX_SERIAL) return E.NUM
    if (!weekend[weekday(d)] && !hol.has(d)) extra--
  }
  return d < 0 || d > MAX_SERIAL ? E.NUM : d
}

export const functions: Record<string, FnDef> = {
  NETWORKDAYS: atp('e e D?', (start: number, end: number, hol?: number[]) => networkDays(start, end, SAT_SUN, hol)),
  'NETWORKDAYS.INTL': atp('e e a? D?', (start: number, end: number, w: V | undefined, hol?: number[]) => {
    const weekend = weekendOf(w)
    return isErr(weekend) ? weekend : networkDays(start, end, weekend, hol)
  }),
  WORKDAY: atp('e j D?', (start: number, days: number, hol?: number[]) => workday(start, days, SAT_SUN, hol)),
  'WORKDAY.INTL': atp('e j a? D?', (start: number, days: number, w: V | undefined, hol?: number[]) => {
    const weekend = weekendOf(w)
    return isErr(weekend) ? weekend : workday(start, days, weekend, hol)
  }),
}
