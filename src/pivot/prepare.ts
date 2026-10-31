/**
 * A PivotTable and its cache made ready to calculate: every field on an axis or the Filters area has its items listed
 * in the cache and on the PivotTable, the records index them, and the value fields' places on the axes are set.
 */
import type { PivotAxisField, PivotCache, PivotField, PivotTable } from '../contract.js'
import { byValue } from './filters.js'
import { groupedItems } from './items.js'
import { withListed } from './items.js'
import { sortLabels } from './label-sort.js'

/** The fields that group the records: on rows, columns or Filters, or filtered by hidden items (a slicer's field) or by label or date (a timeline's). */
function groupingFields(table: PivotTable): Set<number> {
  const out = new Set<number>()
  for (const f of [...(table.rows ?? []), ...(table.columns ?? [])]) if (f !== 'values') out.add(f)
  for (const p of table.pages ?? []) out.add(p.field)
  table.fields.forEach((f, i) => f.items?.some((x) => x.hidden) && out.add(i))
  for (const f of table.filters ?? []) if (!byValue(f)) out.add(f.field)
  return out
}

/**
 * The cache with the fields `need` listed (records copied where one is listed anew), and the PivotTable's fields with
 * an item for each of their cache field's items: those the PivotTable does not list yet come last, shown, in the
 * cache's order, or, for a field it lists none of, sorted as Excel lists a field it places (ascending by label).
 */
export function prepared(table: PivotTable, cache: PivotCache, epoch: number, lists: readonly (readonly string[])[]): { table: PivotTable; cache: PivotCache } {
  const need = groupingFields(table)
  // a grouped field takes its groups' items; groups by hand name the items of their base field, which is listed
  const out = withListed(cache, [...need].map((f) => (cache.fields[f]?.group?.discrete ? (cache.fields[f].group!.base ?? f) : f)))
  const fields = table.fields.map((pf, i): PivotField => {
    const items = groupedItems(out.fields[i])
    if (!need.has(i) || !items) return pf
    const known = new Set((pf.items ?? []).flatMap((x) => (x.item === undefined ? [] : [x.item])))
    const added = items.map((_, k) => k).filter((k) => !known.has(k))
    if (!added.length) return pf
    // groups of a range keep their order, from the first to the last
    if (!pf.items?.length && !out.fields[i].group?.range) sortLabels(added, p => items[p], false, table.customListSort !== false ? lists : null, epoch)
    const calculated = new Set((out.calculatedItems ?? []).flatMap((c) => (c.field === i ? [c.item] : [])))
    return { ...pf, items: [...(pf.items ?? []), ...added.map((item) => (calculated.has(item) ? { item, calculated: true } : { item }))] }
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
