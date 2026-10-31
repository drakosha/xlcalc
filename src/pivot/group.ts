/**
 * A cache field's grouping (`PivotFieldGroup`): each value of its base field to its group, and the groups' labels as
 * Excel makes them: dates by seconds, minutes, hours, days (or a number of days), months, quarters, years; numbers by
 * intervals; items by hand. A value before the start goes to the first group (`<1/1/2024`), after the end to the last
 * (`>12/23/2025`), as does a blank to the first (labelled `(blank)` among numbers).
 */
import type { PivotCache, PivotCacheField, PivotFieldGroup, PivotValue } from '../contract.js'
import { momentOf, ymd } from '../dates.js'
import { generalText } from '../numfmt/general.js'
import { cacheDate, dateSerial, numberOf } from './values.js'

type Range = NonNullable<PivotFieldGroup['range']>

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
/** Days before each month in a leap year: a group by days lists every day of one, 29 February too. */
const BEFORE = [0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335]

const bound = (v: number | { date: string } | undefined, epoch: number): number => (v === undefined ? 0 : typeof v === 'number' ? v : dateSerial(v.date, epoch))

/** A day as Excel's en-US short date (`1/31/2024`). */
function dayText(serial: number): string {
  const { y, m, d } = ymd(serial)
  return `${m}/${d}/${y}`
}

/** A whole number of groups from a quotient that floating point may put a hair above it. */
const groups = (span: number, interval: number) => Math.max(1, Math.ceil(Number((span / interval).toPrecision(12))))

/** The groups between the start and the end of a range: their count, and each value's place among them. */
function inner(range: Range, s: number, e: number): { count: number; place: (x: number) => number } {
  const i = range.interval ?? 1
  const clock = (x: number) => momentOf(x).second
  switch (range.by) {
    case 'numbers': {
      const n = groups(e - s, i)
      return { count: n, place: (x) => Math.min(n - 1, Math.floor(Number(((x - s) / i).toPrecision(12)))) }
    }
    case 'days': {
      if (i > 1) {
        // as many groups as whole intervals reach the end, the last cut short at it (the set group-held: 546 days by 14 are 39)
        const n = groups(Math.floor(e) - Math.floor(s), i)
        return { count: n, place: (x) => Math.min(n - 1, Math.floor((Math.floor(x) - Math.floor(s)) / i)) }
      }
      return { count: 366, place: (x) => { const { m, d } = ymd(x); return BEFORE[m - 1] + d - 1 } }
    }
    case 'months':
      return { count: 12, place: (x) => ymd(x).m - 1 }
    case 'quarters':
      return { count: 4, place: (x) => Math.floor((ymd(x).m - 1) / 3) }
    case 'years': {
      const y0 = ymd(s).y
      return { count: ymd(e).y - y0 + 1, place: (x) => ymd(x).y - y0 }
    }
    case 'hours':
      return { count: 24, place: (x) => Math.floor(clock(x) / 3600) }
    case 'minutes':
      return { count: 60, place: (x) => Math.floor(clock(x) / 60) % 60 }
    case 'seconds':
      return { count: 60, place: (x) => clock(x) % 60 }
  }
}

/** Each value to its group's index in a range grouping of `count` groups with the first and the last. */
export function rangePlace(range: Range, epoch: number): (v: PivotValue) => number {
  const [s, e] = [bound(range.start, epoch), bound(range.end, epoch)]
  const { count, place } = inner(range, s, e)
  return (v) => {
    const x = numberOf(v, epoch)
    if (x === undefined || x < s) return 0
    return x > e ? count + 1 : 1 + place(x)
  }
}

/** A range grouping's labels: the first group (`<start`, `(blank)` among numbers with a blank), those between, the last (`>end`). */
export function rangeLabels(range: Range, blank: boolean, integer: boolean, epoch: number): string[] {
  const [s, e] = [bound(range.start, epoch), bound(range.end, epoch)]
  const { count } = inner(range, s, e)
  const i = range.interval ?? 1
  const g = (x: number) => generalText(Number(x.toPrecision(15)))
  const at = (k: number): string => {
    switch (range.by) {
      case 'numbers': {
        const a = s + k * i
        // whole numbers end a group one below the next's start, the last at the end when that lies on its bound
        const b = integer ? (k === count - 1 ? Math.max(a + i - 1, e) : a + i - 1) : a + i
        return `${g(a)}-${g(b)}`
      }
      case 'days': {
        if (i > 1) {
          const a = Math.floor(s) + k * i
          return `${dayText(a)} - ${dayText(Math.min(a + i - 1, Math.floor(e)))}`
        }
        const m = BEFORE.findLastIndex((b) => b <= k)
        return `${k - BEFORE[m] + 1}-${MONTHS[m]}`
      }
      case 'months':
        return MONTHS[k]
      case 'quarters':
        return `Qtr${k + 1}`
      case 'years':
        return String(ymd(s).y + k)
      case 'hours':
        return `${k % 12 === 0 ? 12 : k % 12} ${k < 12 ? 'AM' : 'PM'}`
      case 'minutes':
      case 'seconds':
        return `:${String(k).padStart(2, '0')}`
    }
  }
  // a blank shares the first group with the values before a start of one's own (the set group-held: `<10 or (blank)`)
  const first = range.by === 'numbers' ? (blank ? (range.autoStart === false ? `<${g(s)} or (blank)` : '(blank)') : `<${g(s)}`) : `<${dayText(s)}`
  const last = range.by === 'numbers' ? `>${g(s + groups(e - s, i) * i)}` : `>${dayText(e)}`
  return [first, ...Array.from({ length: count }, (_, k) => at(k)), last]
}

/**
 * Each record's index into a field's items as the cache lists them: for a grouped field, its group's, by its base
 * field's value (a range) or item (by hand); otherwise the record's own index; −1 where there is none.
 */
export function recordItems(cache: PivotCache, f: number, epoch: number): (r: readonly PivotValue[]) => number {
  const g = cache.fields[f]?.group
  if (!g?.items) return (r) => (typeof r[f] === 'number' ? (r[f] as number) : -1)
  const b = g.base ?? f
  const baseItems = cache.fields[b]?.items
  if (g.discrete) {
    const groups = g.discrete
    return (r) => (typeof r[b] === 'number' ? (groups[r[b] as number] ?? -1) : -1)
  }
  if (!g.range) return () => -1
  const place = rangePlace(g.range, epoch)
  return (r) => {
    const raw = r[b] ?? null
    return place(baseItems && typeof raw === 'number' ? (baseItems[raw] ?? null) : raw)
  }
}

const key = (v: PivotValue) => JSON.stringify(v)

/**
 * The groups of a refreshed cache's fields: base and parent moved with their fields (`fieldMap`, old index to new), a
 * range's automatic start and end taken again from its base field's values (a last date of whole days a day on, as
 * Excel takes it) and its labels made again, groups by hand moved with their base field's items (`itemMaps`), an item
 * new to it a group of its own. Gives, for each grouped field, its old groups' new indexes by label.
 */
export function regroup(fields: PivotCacheField[], fieldMap: ReadonlyMap<number, number>, itemMaps: readonly (ReadonlyMap<number, number> | undefined)[], epoch: number): Map<number, Map<number, number>> {
  const out = new Map<number, Map<number, number>>()
  fields.forEach((f, n) => {
    const was = f.group
    if (!was) return
    const g: PivotFieldGroup = { ...was }
    const to = (i: number | undefined) => (i === undefined ? undefined : fieldMap.get(i))
    if (was.base !== undefined) g.base = to(was.base)
    if (was.parent !== undefined) g.parent = to(was.parent)
    for (const k of ['base', 'parent'] as const) if (g[k] === undefined) delete g[k]
    const b = g.base ?? n
    const base = fields[b]
    if (was.base !== undefined && g.base === undefined) {
      // its base field is gone from the source: the groups stay as they were
      fields[n] = { ...f, group: g }
      return
    }
    if (g.range && was.items && base?.facts) {
      const facts = base.facts
      const range = { ...g.range }
      if (range.autoStart !== false && facts.min !== undefined) range.start = facts.min
      if (range.autoEnd !== false && facts.max !== undefined) {
        const max = facts.max
        range.end = typeof max === 'object' && max.date.endsWith('T00:00:00') ? { date: cacheDate(dateSerial(max.date, epoch) + 1, epoch) } : max
      }
      g.range = range
      g.items = rangeLabels(range, facts.types.includes('blank'), facts.integer === true, epoch)
      const place = new Map(g.items.map((v, i) => [key(v), i]))
      out.set(n, new Map(was.items.flatMap((v, i) => (place.has(key(v)) ? [[i, place.get(key(v))!]] : []))))
    } else if (g.discrete && g.items && base?.items) {
      const moved = itemMaps[b] ?? new Map(was.discrete!.map((_, i) => [i, i]))
      const discrete = new Array<number>(base.items.length).fill(-1)
      was.discrete!.forEach((grp, i) => {
        const j = moved.get(i)
        if (j !== undefined) discrete[j] = grp
      })
      const items = g.items.slice()
      discrete.forEach((grp, j) => grp < 0 && (discrete[j] = items.push(base.items![j]) - 1))
      g.discrete = discrete
      g.items = items
      out.set(n, new Map(was.items!.map((_, i) => [i, i])))
    }
    fields[n] = { ...f, group: g }
  })
  return out
}
