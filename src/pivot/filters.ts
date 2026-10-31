/**
 * A PivotTable's filters of a field: by label and by date, which test each item alone, and by value and Top 10, which
 * test an item's value among the items under the same item of the fields before it.
 */
import { collator } from '../collate.js'
import type { CellValue, PivotFilter, PivotValue } from '../contract.js'
import { ymd } from '../dates.js'
import { Wildcard } from '../wildcards.js'
import { dateSerial, itemText } from './values.js'

const LABEL = /^caption/
const VALUE = /^value/
const TOP = new Set(['count', 'percent', 'sum'])

/** A filter that tests an item by its value field's value (value filters, Top 10). */
export const byValue = (f: PivotFilter): boolean => VALUE.test(f.type) || TOP.has(f.type)

const wild = (pattern: string, text: string) => new Wildcard(pattern, 'iu').test(text)

/** Two texts, or two numbers, in Excel's order. */
function compared(text: string, item: PivotValue, arg: string | number): number {
  if (typeof item === 'number' && typeof arg === 'number') return item - arg
  return collator.compare(text, String(arg))
}

function label(f: PivotFilter, item: PivotValue, caption: string | undefined, epoch: number): boolean {
  const text = caption ?? itemText(item, epoch)
  const [a = '', b = ''] = f.args ?? []
  const s = String(a)
  switch (f.type.replace(/^captionNot/, 'caption')) {
    case 'captionEqual':
      return wild(s, text) !== (f.type === 'captionNotEqual')
    case 'captionBeginsWith':
      return wild(`${s}*`, text) !== (f.type === 'captionNotBeginsWith')
    case 'captionEndsWith':
      return wild(`*${s}`, text) !== (f.type === 'captionNotEndsWith')
    case 'captionContains':
      return wild(`*${s}*`, text) !== (f.type === 'captionNotContains')
    case 'captionGreaterThan':
      return compared(text, item, a) > 0
    case 'captionGreaterThanOrEqual':
      return compared(text, item, a) >= 0
    case 'captionLessThan':
      return compared(text, item, a) < 0
    case 'captionLessThanOrEqual':
      return compared(text, item, a) <= 0
    case 'captionBetween': {
      const inside = compared(text, item, a) >= 0 && compared(text, item, b) <= 0
      return inside !== (f.type === 'captionNotBetween')
    }
  }
  return true
}

const quarter = (m: number) => Math.floor((m - 1) / 3)

/** Whether a date (a serial) passes a date filter, relative ones against `today`. */
function date(f: PivotFilter, x: number, today: number, epoch: number): boolean {
  const whole = f.wholeDays !== false
  const at = (k: number) => {
    const arg = f.args?.[k] ?? 0
    const v = typeof arg === 'number' ? arg : dateSerial(arg, epoch)
    return whole ? Math.floor(v) : v
  }
  const d = Math.floor(x)
  const v = whole ? d : x
  const t = Math.floor(today)
  const { y, m } = ymd(d + epoch)
  const now = ymd(t + epoch)
  const months = (y - now.y) * 12 + m - now.m
  const quarters = (y - now.y) * 4 + quarter(m) - quarter(now.m)
  // weeks from Sunday: serial 1 (1900-01-01) is a Sunday
  const week = (s: number) => Math.floor((s + epoch - 1) / 7)
  const type = f.type
  switch (type) {
    case 'dateEqual':
      return v === at(0)
    case 'dateNotEqual':
      return v !== at(0)
    case 'dateOlderThan':
      return v < at(0)
    case 'dateOlderThanOrEqual':
      return v <= at(0)
    case 'dateNewerThan':
      return v > at(0)
    case 'dateNewerThanOrEqual':
      return v >= at(0)
    case 'dateBetween':
    case 'dateNotBetween':
      return (v >= at(0) && v <= at(1)) !== (type === 'dateNotBetween')
    case 'yesterday':
    case 'today':
    case 'tomorrow':
      return d - t === ['yesterday', 'today', 'tomorrow'].indexOf(type) - 1
    case 'lastWeek':
    case 'thisWeek':
    case 'nextWeek':
      return week(d) - week(t) === ['lastWeek', 'thisWeek', 'nextWeek'].indexOf(type) - 1
    case 'lastMonth':
    case 'thisMonth':
    case 'nextMonth':
      return months === ['lastMonth', 'thisMonth', 'nextMonth'].indexOf(type) - 1
    case 'lastQuarter':
    case 'thisQuarter':
    case 'nextQuarter':
      return quarters === ['lastQuarter', 'thisQuarter', 'nextQuarter'].indexOf(type) - 1
    case 'lastYear':
    case 'thisYear':
    case 'nextYear':
      return y - now.y === ['lastYear', 'thisYear', 'nextYear'].indexOf(type) - 1
    case 'yearToDate':
      return y === now.y && d <= t
  }
  if (/^Q[1-4]$/.test(type)) return quarter(m) === Number(type[1]) - 1
  if (/^M\d+$/.test(type)) return m === Number(type.slice(1))
  return true
}

/**
 * Whether an item passes a filter that tests it alone: a label filter by its caption or text, a date filter by its date
 * (an item of no date fails one). A value filter or Top 10 passes here.
 */
export function itemPasses(f: PivotFilter, item: PivotValue, caption: string | undefined, today: number, epoch: number): boolean {
  if (LABEL.test(f.type)) return label(f, item, caption, epoch)
  if (byValue(f)) return true
  if (item === null || typeof item !== 'object' || !('date' in item)) return false
  return date(f, dateSerial(item.date, epoch), today, epoch)
}

/**
 * The items of one group (those under the same items of the fields before) that a value filter or Top 10 keeps, by
 * their values: an empty value counts as 0, an error fails them. Top 10 by count keeps the items tied with the last one it
 * takes; by percent or sum it takes no more once it has enough, of tied items the earlier first (the set filter-held).
 */
export function valuesKept(f: PivotFilter, values: readonly CellValue[]): boolean[] {
  // an empty value is 0 to them (the set filter: (blank) kept by "less than"), an error fails them
  const nums = values.map((v) => (typeof v === 'number' ? v : v === null ? 0 : undefined))
  const [a = 0, b = 0] = (f.args ?? []).map(Number)
  if (!TOP.has(f.type)) {
    const [lo, hi] = a <= b ? [a, b] : [b, a]
    const test: Record<string, (x: number) => boolean> = {
      valueEqual: (x) => x === a,
      valueNotEqual: (x) => x !== a,
      valueGreaterThan: (x) => x > a,
      valueGreaterThanOrEqual: (x) => x >= a,
      valueLessThan: (x) => x < a,
      valueLessThanOrEqual: (x) => x <= a,
      valueBetween: (x) => x >= lo && x <= hi,
      valueNotBetween: (x) => x < lo || x > hi,
    }
    return nums.map((x) => x !== undefined && (test[f.type]?.(x) ?? true))
  }
  const order = nums.flatMap((x, i) => (x === undefined ? [] : [i])).sort((i, j) => (f.bottom ? nums[i]! - nums[j]! : nums[j]! - nums[i]!) || i - j)
  const kept = values.map(() => false)
  const total = order.reduce((s, i) => s + nums[i]!, 0)
  const target = f.type === 'percent' ? (total * a) / 100 : a
  let taken = 0,
    sum = 0,
    last: number | undefined
  for (const i of order) {
    // by count, the items tied with the last one taken come too; by percent or sum, no more once there is enough
    if (f.type === 'count' ? taken >= a && nums[i] !== last : taken > 0 && sum >= target) break
    kept[i] = true
    taken++
    sum += nums[i]!
    last = nums[i]
  }
  return kept
}
