/**
 * A PivotTable and its cache made ready to calculate: every field on an axis or the Filters area has its items listed
 * in the cache and on the PivotTable, the records index them, and the value fields' places on the axes are set.
 */
import type { PivotAxisField, PivotCache, PivotCacheField, PivotField, PivotTable, PivotValue } from '../contract.js'
import { itemKey, labelOrder } from './values.js'

/** The fields that group the records: on rows, columns or Filters, or filtered by hidden items (a slicer's field). */
function groupingFields(table: PivotTable): Set<number> {
  const out = new Set<number>()
  for (const f of [...(table.rows ?? []), ...(table.columns ?? [])]) if (f !== 'values') out.add(f)
  for (const p of table.pages ?? []) out.add(p.field)
  table.fields.forEach((f, i) => f.items?.some((x) => x.hidden) && out.add(i))
  return out
}

/** A field's values listed as items in the order they first come, the records turned to their indexes. */
function listed(field: PivotCacheField, records: PivotValue[][], at: number): PivotCacheField {
  const items: PivotValue[] = []
  const seen = new Map<string, number>()
  for (const r of records) {
    const v = r[at] ?? null
    const k = itemKey(v)
    let i = seen.get(k)
    if (i === undefined) {
      i = items.length
      seen.set(k, i)
      items.push(v)
    }
    r[at] = i
  }
  return { ...field, items }
}

/**
 * The cache with the fields `need` listed (records copied where one is listed anew), and the PivotTable's fields with
 * an item for each of their cache field's items: those the PivotTable does not list yet come last, shown, in the
 * cache's order, or, for a field it lists none of, sorted as Excel lists a field it places (ascending by label).
 */
export function prepared(table: PivotTable, cache: PivotCache, epoch: number, lists: readonly (readonly string[])[]): { table: PivotTable; cache: PivotCache } {
  const need = groupingFields(table)
  const unlisted = [...need].filter((f) => cache.fields[f] && !cache.fields[f].items)
  let out = cache
  if (unlisted.length) {
    const records = (cache.records ?? []).map((r) => r.slice())
    const fields = cache.fields.slice()
    for (const f of unlisted) fields[f] = listed(fields[f], records, f)
    out = { ...cache, fields, records }
  }
  const fields = table.fields.map((pf, i): PivotField => {
    const items = out.fields[i]?.items
    if (!need.has(i) || !items) return pf
    const known = new Set((pf.items ?? []).flatMap((x) => (x.item === undefined ? [] : [x.item])))
    const added = items.map((_, k) => k).filter((k) => !known.has(k))
    if (!added.length) return pf
    if (!pf.items?.length) added.sort((a, b) => labelOrder(items[a], items[b], false, table.customListSort !== false ? lists : null, epoch) || a - b)
    return { ...pf, items: [...(pf.items ?? []), ...added.map((item) => ({ item }))] }
  })
  return { table: { ...table, fields }, cache: out }
}

/**
 * The fields of the rows and of the columns, with `values` where the value fields lie when there are several: where an
 * axis names it, else at `valuesPosition` (last without it) of the rows (`valuesOnRows`) or of the columns.
 */
export function axesOf(table: PivotTable): { rows: PivotAxisField[]; columns: PivotAxisField[] } {
  const rows = table.rows ?? []
  const columns = table.columns ?? []
  const named = rows.includes('values') || columns.includes('values')
  if (named || (table.values?.length ?? 0) < 2) return { rows: rows.slice(), columns: columns.slice() }
  const axis = (table.valuesOnRows ? rows : columns).slice()
  axis.splice(Math.min(table.valuesPosition ?? axis.length, axis.length), 0, 'values')
  return table.valuesOnRows ? { rows: axis, columns: columns.slice() } : { rows: rows.slice(), columns: axis }
}
