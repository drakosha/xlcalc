/** The number format a PivotTable shows an item's label in, which GETPIVOTDATA matches text against and the layout writes; a slicer's labels of its items. */
import type { BookInput, NumberFormat, PivotCache, PivotValue, SlicerCache, SlicerCacheLayout } from './contract.js'
import { dateToSerial } from './date-serial.js'
import { epochOf } from './dates.js'
import { displayRules } from './locale/rules.js'
import { groupedItems } from './pivot/items.js'
import { blankWord } from './pivot-words.js'
import { caseKey } from './collate.js'
import { sheetKey } from './syntax.js'

/**
 * An item's format: its cache field's own; else a date's the short date, a date with a time the short date and time
 * (built-in 14 and 22), as Excel shows the items of a field of dates in the General format (the sets basic and gpd-1).
 */
export function itemFormat(c: PivotCache, f: number, v: PivotValue): NumberFormat | undefined {
  const own = c.fields[f]?.numberFormat
  if (own !== undefined || v === null || typeof v !== 'object' || !('date' in v)) return own
  return v.date.endsWith('T00:00:00') || !v.date.includes('T') ? 14 : 22
}

/** A number as a slicer names it without a format: its 15 significant digits written out, no exponent (`100000000000000000000`, `0.333333333333333`). */
export function slicerNumber(n: number, decimal: string): string {
  const [digits, exp] = n.toPrecision(15).split('e')
  const shift = Number(exp ?? 0)
  const negative = digits.startsWith('-')
  const [whole, fraction = ''] = digits.replace('-', '').split('.')
  const all = whole + fraction
  // the decimal point moved by the exponent, zeros added on either side
  const point = whole.length + shift
  const int = point <= 0 ? '0' : all.slice(0, point).padEnd(point, '0')
  const frac = (point < 0 ? '0'.repeat(-point) + all : all.slice(Math.max(0, point))).replace(/0+$/, '')
  return `${negative ? '-' : ''}${int.replace(/^0+(?=\d)/, '')}${frac ? decimal + frac : ''}`
}

/**
 * A slicer cache with each item's label as the slicer shows it in the workbook's locale: its value in the number
 * format of its field on the first PivotTable the cache filters, else its cache field's (`itemFormat`), a number in
 * none written out (`slicerNumber`); a blank in the locale's word, a boolean and an error in the locale's words, a
 * group's name; empty for an item its pivot cache does not hold (the set slicer-labels).
 */
export function labelled(book: Pick<BookInput, 'locale' | 'dateSystem' | 'pivotCaches' | 'sheets'>, sc: SlicerCache): SlicerCacheLayout {
  if (!sc.items) return sc as SlicerCacheLayout
  const pc = sc.cache === undefined ? undefined : book.pivotCaches?.find((c) => c.id === sc.cache)
  const f = pc ? pc.fields.findIndex((x) => caseKey(x.name) === caseKey(sc.field)) : -1
  const values = f >= 0 ? groupedItems(pc!.fields[f]) : undefined
  const first = sc.pivots[0] && book.sheets.find((s) => sheetKey(s.name) === sheetKey(sc.pivots[0].sheet))?.pivotTables?.find((p) => caseKey(p.name) === caseKey(sc.pivots[0].name))
  const own = first?.fields[f]?.numberFormat
  const rules = displayRules(book.locale, epochOf(book.dateSystem))
  const label = (k: number): string => {
    if (!values || k >= values.length) return ''
    const v = values[k]
    if (v === null) return blankWord(book.locale)
    if (typeof v === 'string') return v
    // a boolean and an error are the locale's words for them, in no format
    if (typeof v === 'boolean' || (typeof v === 'object' && 'error' in v)) return rules.display(v, undefined, Infinity).text
    const format = own ?? itemFormat(pc!, f, v)
    if (typeof v === 'number' && format === undefined) return slicerNumber(v, rules.decimal)
    return rules.display(typeof v === 'object' ? dateToSerial(v.date, { dateSystem: book.dateSystem }) : v, format, Infinity).text
  }
  return { ...sc, items: sc.items.map((x) => ({ ...x, label: label(x.item) })) }
}
