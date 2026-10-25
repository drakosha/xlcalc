/** Workdays: NETWORKDAYS and WORKDAY with weekend codes and holidays. */
import { eachValue, toDate, type Cx, type FnDef, type Segment } from '../../args.js'
import type { Rules } from '../../locale/rules.js'
import { MAX_SERIAL } from '../../dates.js'
import { E, isErr, type CErr, type V } from '../../values.js'
import { atp } from './atp.js'
import { weekday } from './calendar.js'

/** Weekend mask (index 0 Sunday) by code 1–7, 11–17 or seven 0/1 from Monday; Excel answers an error there with #NUM!. */
function weekendOf(w: V | undefined): boolean[] | CErr {
  const mask = [false, false, false, false, false, false, false]
  if (w === undefined) return weekendOf(1)
  if (isErr(w)) return E.NUM
  if (typeof w === 'string') {
    if (!/^[01]{7}$/.test(w)) return E.VALUE
    for (let i = 0; i < 7; i++) mask[(i + 1) % 7] = w[i] === '1'
    return mask
  }
  const k = Math.trunc(Number(w))
  if (k >= 1 && k <= 7) {
    mask[(k + 5) % 7] = true
    mask[(k + 6) % 7] = true
  } else if (k >= 11 && k <= 17) mask[k - 11] = true
  else return E.NUM
  return mask
}

const SAT_SUN = weekendOf(1) as boolean[]
const allWeekend = (w: boolean[]) => w.every(Boolean)

/** Holidays: whole dates, date text in a range too (the set `findate-edges`); a boolean is #VALUE! (the ToolPak rule), so is other text. */
function holidaysOf(x: Segment | undefined, rules: Rules): number[] | CErr {
  const out: number[] = []
  const e = x === undefined ? undefined : eachValue([x], (v, fromRange) => {
    // a blank cell of a holidays range is no holiday; a blank given as a value is a date
    if (v === null && fromRange) return
    if (isErr(v)) return v
    if (typeof v === 'boolean') return E.VALUE
    const d = toDate(v, rules)
    if (isErr(d)) return d
    out.push(d)
  })
  return e ?? out
}

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
  NETWORKDAYS: atp('e e V?', (cx: Cx, start: number, end: number, h?: Segment) => {
    const hol = holidaysOf(h, cx.rules)
    return isErr(hol) ? hol : networkDays(start, end, SAT_SUN, hol)
  }, { cx: true }),
  // seven weekend days leave no workdays: 0 here, #VALUE! in WORKDAY.INTL
  'NETWORKDAYS.INTL': atp('e e a? V?', (cx: Cx, start: number, end: number, w: V | undefined, h?: Segment) => {
    const weekend = weekendOf(w)
    if (isErr(weekend)) return weekend
    const hol = holidaysOf(h, cx.rules)
    if (isErr(hol)) return hol
    return allWeekend(weekend) ? 0 : networkDays(start, end, weekend, hol)
  }, { cx: true }),
  WORKDAY: atp('e f V?', (cx: Cx, start: number, days: number, h?: Segment) => {
    const hol = holidaysOf(h, cx.rules)
    return isErr(hol) ? hol : workday(start, days, SAT_SUN, hol)
  }, { cx: true, date: 'book' }),
  'WORKDAY.INTL': atp('e f a? V?', (cx: Cx, start: number, days: number, w: V | undefined, h?: Segment) => {
    const weekend = weekendOf(w)
    if (isErr(weekend)) return weekend
    if (allWeekend(weekend)) return E.VALUE
    const hol = holidaysOf(h, cx.rules)
    return isErr(hol) ? hol : workday(start, days, weekend, hol)
  }, { cx: true, date: 'book' }),
}
