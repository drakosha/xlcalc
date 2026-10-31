/** The number format a PivotTable shows an item's label in, which GETPIVOTDATA matches text against and the layout writes. */
import type { NumberFormat, PivotCache, PivotValue } from './contract.js'

/**
 * An item's format: its cache field's own; else a date's the short date, a date with a time the short date and time
 * (built-in 14 and 22), as Excel shows the items of a field of dates in the General format (the sets basic and gpd-1).
 */
export function itemFormat(c: PivotCache, f: number, v: PivotValue): NumberFormat | undefined {
  const own = c.fields[f]?.numberFormat
  if (own !== undefined || v === null || typeof v !== 'object' || !('date' in v)) return own
  return v.date.endsWith('T00:00:00') || !v.date.includes('T') ? 14 : 22
}
