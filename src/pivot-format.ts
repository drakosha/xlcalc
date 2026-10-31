/** The number format a PivotTable shows an item's label in, which GETPIVOTDATA matches text against and the layout writes; a slicer's labels of its items. */
import type { BookInput, NumberFormat, PivotCache, PivotValue, SlicerCache, SlicerCacheLayout } from './contract.js'
import { dateToSerial } from './date-serial.js'
import { epochOf } from './dates.js'
import { displayRules } from './locale/rules.js'
import { groupedItems } from './pivot/items.js'
import { blankWord } from './pivot-words.js'
import { caseKey } from './collate.js'

/**
 * An item's format: its cache field's own; else a date's the short date, a date with a time the short date and time
 * (built-in 14 and 22), as Excel shows the items of a field of dates in the General format (the sets basic and gpd-1).
 */
export function itemFormat(c: PivotCache, f: number, v: PivotValue): NumberFormat | undefined {
  const own = c.fields[f]?.numberFormat
  if (own !== undefined || v === null || typeof v !== 'object' || !('date' in v)) return own
  return v.date.endsWith('T00:00:00') || !v.date.includes('T') ? 14 : 22
}

/**
 * A slicer cache with each item's label as the slicer shows it in the workbook's locale: its value in its field's
 * format (`itemFormat`), a blank in the locale's word, a boolean and an error in the locale's words, a group's name;
 * empty for an item its pivot cache does not hold.
 */
export function labelled(book: Pick<BookInput, 'locale' | 'dateSystem' | 'pivotCaches'>, sc: SlicerCache): SlicerCacheLayout {
  if (!sc.items) return sc as SlicerCacheLayout
  const pc = sc.cache === undefined ? undefined : book.pivotCaches?.find((c) => c.id === sc.cache)
  const f = pc ? pc.fields.findIndex((x) => caseKey(x.name) === caseKey(sc.field)) : -1
  const values = f >= 0 ? groupedItems(pc!.fields[f]) : undefined
  const rules = displayRules(book.locale, epochOf(book.dateSystem))
  const label = (k: number): string => {
    if (!values || k >= values.length) return ''
    const v = values[k]
    if (v === null) return blankWord(book.locale)
    if (typeof v === 'string') return v
    // a boolean and an error are the locale's words for them, in no format
    if (typeof v === 'boolean' || (typeof v === 'object' && 'error' in v)) return rules.display(v, undefined, Infinity).text
    const shown = typeof v === 'object' ? dateToSerial(v.date, { dateSystem: book.dateSystem }) : v
    return rules.display(shown, itemFormat(pc!, f, v), Infinity).text
  }
  return { ...sc, items: sc.items.map((x) => ({ ...x, label: label(x.item) })) }
}
