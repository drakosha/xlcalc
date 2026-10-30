/** A pivot cache's values as the calculation takes them: a key per item, a number to summarize, a caption, an order. */
import { caseKey } from '../collate.js'
import type { PivotValue } from '../contract.js'
import { ymd, momentOf } from '../dates.js'
import { dateToSerial } from '../model.js'
import { generalText } from '../numfmt/general.js'
import { sortOrder } from '../sort-order.js'
import { ERROR_TEXT } from '../syntax/errors.js'

/** A date of a cache as a serial number of the workbook's date system (`epoch` 1462 for 1904). */
export const dateSerial = (iso: string, epoch = 0): number => dateToSerial(iso) - epoch

const pad = (n: number, w = 2) => String(n).padStart(w, '0')

/** A serial number as a cache saves a date: ISO 8601 to the second, without a zone (`2025-01-31T00:00:00`). */
export function cacheDate(serial: number, epoch = 0): string {
  const { day, second } = momentOf(serial + epoch)
  const { y, m, d } = ymd(day)
  return `${pad(y, 4)}-${pad(m)}-${pad(d)}T${pad(Math.floor(second / 3600))}:${pad(Math.floor(second / 60) % 60)}:${pad(second % 60)}`
}

/**
 * The key two values are one item by: text in any case is one (`North` and `north`), a date is no number, each error
 * its own.
 */
export function itemKey(v: PivotValue): string {
  if (v === null) return 'b'
  switch (typeof v) {
    case 'number':
      return `n${v}`
    case 'string':
      return `s${caseKey(v)}`
    case 'boolean':
      return v ? 'T' : 'F'
  }
  return 'error' in v ? `e${v.error}` : `d${v.date}`
}

/** A value as a number to summarize: a number, or a date's serial; nothing else counts as one. */
export const numberOf = (v: PivotValue, epoch = 0): number | undefined =>
  typeof v === 'number' ? v : v !== null && typeof v === 'object' && 'date' in v ? dateSerial(v.date, epoch) : undefined

/** A value as the calculation compares it: a date as its serial, an error as `{ error }`, the rest as is. */
export const plainOf = (v: PivotValue, epoch = 0): number | string | boolean | null | { error: string } =>
  v !== null && typeof v === 'object' && 'date' in v ? dateSerial(v.date, epoch) : v

/** An item's text when no caption replaces it, which label filters compare: `(blank)` for a blank. */
export function itemText(v: PivotValue, epoch = 0): string {
  if (v === null) return '(blank)'
  switch (typeof v) {
    case 'number':
      return generalText(v)
    case 'string':
      return v
    case 'boolean':
      return v ? 'TRUE' : 'FALSE'
  }
  return 'error' in v ? ERROR_TEXT[v.error] : generalText(dateSerial(v.date, epoch))
}

/** Custom lists by a text's key: its place, list by list, then within one. */
const places = new WeakMap<readonly (readonly string[])[], Map<string, number>>()
function listPlace(lists: readonly (readonly string[])[], s: string): number | undefined {
  let m = places.get(lists)
  if (!m) places.set(lists, (m = new Map(lists.flatMap((l, i) => l.map((w, j): [string, number] => [caseKey(w), i * 1000 + j])).reverse())))
  return m.get(caseKey(s))
}

/**
 * The order of items by label: the texts of a custom list first, by its order (with `lists`, the first list a text is on), then numbers and dates,
 * text, booleans, errors, the blank last either way.
 */
export function labelOrder(a: PivotValue, b: PivotValue, desc: boolean, lists: readonly (readonly string[])[] | null, epoch = 0): number {
  if (lists && typeof a === 'string' && typeof b === 'string') {
    const [x, y] = [listPlace(lists, a), listPlace(lists, b)]
    if (x !== undefined || y !== undefined) {
      const c = x === undefined ? 1 : y === undefined ? -1 : x - y
      return desc ? -c : c
    }
  }
  return sortOrder(plainOf(a, epoch), plainOf(b, epoch), desc)
}
