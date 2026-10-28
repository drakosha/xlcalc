/** Workdays: NETWORKDAYS and WORKDAY with weekend codes and holidays. */
import { eachValue, toDate, type Cx, type FnDef, type Segment } from '../../args.js'
import type { Rules } from '../../locale/rules.js'
import { MAX_SERIAL } from '../../dates.js'
import { E, isErr, type CErr, type V } from '../../values.js'
import { readsClock } from '../../locale/rules/data.js'
import { keptWith, VALUE_BYTES } from '../../range-cache.js'
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
function holidayList(x: Segment | undefined, rules: Rules): Float64Array | CErr {
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
  if (e) return e
  // ascending and once each: the counts below take a holiday once, by binary search
  const sorted = Float64Array.from(out).sort()
  let n = 0
  for (let i = 0; i < sorted.length; i++) if (i === 0 || sorted[i] !== sorted[n - 1]) sorted[n++] = sorted[i]
  return sorted.subarray(0, n)
}

// Range arrays are immutable and replaced on invalidation: a long holidays range filled down a column is read once.
const kept = new WeakMap<readonly V[], { rules: Rules; holidays: Float64Array | CErr }>()
const KEPT_LENGTH = 256

/** `holidayList`, kept with a long range's values unless reading them took the clock's year (date text without one). */
function holidaysOf(x: Segment | undefined, rules: Rules): Float64Array | CErr {
  if (!Array.isArray(x) || x.length < KEPT_LENGTH) return holidayList(x, rules)
  const hit = kept.get(x)
  if (hit?.rules === rules) return hit.holidays
  const [holidays, clocked] = readsClock(rules.clock, () => holidayList(x, rules))
  if (!clocked) {
    kept.set(x, { rules, holidays })
    keptWith(x, 64 + (isErr(holidays) ? 0 : holidays.length * VALUE_BYTES))
  }
  return holidays
}

/** The first position in ascending `xs` whose value is above `x`, or at least `x` when `equal`. */
function bound(xs: Float64Array, x: number, equal: boolean): number {
  let lo = 0, hi = xs.length
  while (lo < hi) {
    const mid = (lo + hi) >>> 1
    if (xs[mid] < x || (!equal && xs[mid] === x)) lo = mid + 1
    else hi = mid
  }
  return lo
}

/** Holidays on workdays from position `i` on, up to `b` (inclusive when `upTo`). */
function workdayHolidays(holidays: Float64Array, i: number, b: number, upTo: boolean, weekend: boolean[]): number {
  let n = 0
  for (; i < holidays.length && (holidays[i] < b || (upTo && holidays[i] === b)); i++) if (!weekend[weekday(holidays[i])]) n++
  return n
}

const EMPTY: Float64Array = new Float64Array(0)

/** Inclusive count, negative when `end` is earlier. */
function networkDays(start: number, end: number, weekend: boolean[], holidays = EMPTY): number {
  const sign = start <= end ? 1 : -1
  const [a, b] = sign > 0 ? [start, end] : [end, start]
  const weeks = Math.floor((b - a + 1) / 7)
  let count = weeks * weekend.filter((x) => !x).length
  for (let s = a + weeks * 7; s <= b; s++) if (!weekend[weekday(s)]) count++
  return sign * (count - workdayHolidays(holidays, bound(holidays, a, true), b, true, weekend))
}

/** Whole weeks first, then day by day; holidays passed on workdays are made up past the date reached. */
function workday(start: number, days: number, weekend: boolean[], holidays = EMPTY): number | CErr {
  if (days === 0) return start
  if (Math.abs(days) > MAX_SERIAL) return E.NUM
  const step = days > 0 ? 1 : -1
  const perWeek = weekend.filter((x) => !x).length
  const weeks = Math.floor((Math.abs(days) - 1) / perWeek)
  let left = Math.abs(days) - weeks * perWeek
  let d = start + step * weeks * 7
  while (left > 0) {
    d += step
    if (!weekend[weekday(d)]) left--
  }
  // after `start` up to `d`, or from `d` up to before `start`
  let extra = step > 0 ? workdayHolidays(holidays, bound(holidays, start, false), d, true, weekend) : workdayHolidays(holidays, bound(holidays, d, true), start, false, weekend)
  const holiday = (x: number) => { const i = bound(holidays, x, true); return i < holidays.length && holidays[i] === x }
  while (extra > 0) {
    d += step
    if (d < 0 || d > MAX_SERIAL) return E.NUM
    if (!weekend[weekday(d)] && !holiday(d)) extra--
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
