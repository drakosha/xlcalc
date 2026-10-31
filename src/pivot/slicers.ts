/**
 * A slicer's items as Excel shows them after its PivotTables are calculated: chosen unless hidden on the first
 * PivotTable it filters, marked when no record has them under the other fields' filters, in its order (by label, the
 * custom lists first, ascending by default) with the items with data first unless its cache says otherwise.
 */
import type { BookInput, PivotCache, PivotRef, PivotTable, PivotValue, SlicerCache, TimelineCache } from '../contract.js'
import { groupedItems } from './items.js'
import { ERROR_TYPE } from '../error-type.js'
import { caseKey } from '../collate.js'
import { fieldIndex, sameRef, unchosen } from '../structure/slicer-choice.js'
import { itemsWithData } from './compute.js'
import type { PivotComputeOptions } from './result.js'
import { labelOrder } from './values.js'
import { EPOCH_1904 } from '../dates.js'

/** Errors in a slicer's ascending order: by `ERROR.TYPE` down, `#N/A`, `#NUM!`, `#DIV/0!`, as Excel's PivotTable lists them too (the set slicer-labels). */
function errorOrder(a: PivotValue, b: PivotValue, desc: boolean): number {
  const type = (v: PivotValue) => (v !== null && typeof v === 'object' && 'error' in v ? (ERROR_TYPE[v.error] ?? 0) : 0)
  const c = type(b) - type(a)
  return desc ? -c : c
}

/** A slicer cache's items from the calculations of its PivotTables (`results`, the first the one it reads its choice from); the cache as it was without one. */
export function slicerItems(sc: SlicerCache, results: readonly { table: PivotTable; cache: PivotCache }[], options: PivotComputeOptions): SlicerCache {
  const first = results[0]
  const f = first ? fieldIndex(first.cache, sc.field) : -1
  // a grouped field's slicer lists its groups (the set slicer-labels)
  const values = f >= 0 ? groupedItems(first.cache.fields[f]) : undefined
  if (!values) return sc
  const data = new Set<number>()
  for (const r of results) for (const k of itemsWithData(r.table, r.cache, f, options)) data.add(k)
  const hidden = unchosen(first.table, f)
  const epoch = options.date1904 ? EPOCH_1904 : 0
  const lists = sc.customListSort !== false ? (options.lists ?? null) : null
  const order = values.map((_, k) => k).sort((a, b) => labelOrder(values[a], values[b], sc.sort === 'descending', lists, epoch) || errorOrder(values[a], values[b], sc.sort === 'descending') || a - b)
  // Excel's default: the items with data at the top
  if (sc.noData?.mark !== false && sc.noData?.last !== false) order.sort((a, b) => Number(!data.has(a)) - Number(!data.has(b)))
  // without cross filtering (`crossFilter="none"`) Excel marks no item as without data
  const mark = sc.noData?.mark !== false
  return { ...sc, items: order.map((item) => ({ item, ...(hidden.has(item) ? { selected: false } : {}), ...(!mark || data.has(item) ? {} : { noData: true }) })) }
}

/**
 * A PivotTable as Excel calculates it: a field on no axis filters it only while a slicer or timeline over it names the
 * field. Its hidden items and its label and date filters are otherwise what a deleted timeline or a disconnected slicer
 * left in the file, which Excel saves and no longer applies, also on a refresh (the sets slicer-file and slicer-held).
 * Excel filters by label or date only a field on rows or columns, and a timeline's field anywhere.
 */
export function liveFilters(table: PivotTable, ref: PivotRef, book: Pick<BookInput, 'slicerCaches' | 'timelineCaches'>, fields: readonly { name: string }[]): PivotTable {
  const placed = new Set<number | 'values'>([...(table.rows ?? []), ...(table.columns ?? []), ...(table.pages ?? []).map((p) => p.field)])
  const over = (caches: readonly (SlicerCache | TimelineCache)[] | undefined) => new Set((caches ?? []).filter((c) => c.pivots.some((r) => sameRef(r, ref))).map((c) => caseKey(c.field)))
  const [sliced, timed] = [over(book.slicerCaches), over(book.timelineCaches)]
  const named = (f: number, by: Set<string>) => placed.has(f) || by.has(caseKey(fields[f]?.name ?? ''))
  const stale = table.fields.some((pf, f) => !named(f, sliced) && pf.items?.some((x) => x.hidden))
  const filters = table.filters?.filter((x) => named(x.field, timed))
  if (!stale && filters?.length === table.filters?.length) return table
  const { filters: _, ...rest } = table
  const shown = stale ? table.fields.map((pf, f) => (named(f, sliced) || !pf.items ? pf : { ...pf, items: pf.items.map(({ hidden: __, ...x }) => x) })) : table.fields
  return { ...rest, fields: shown, ...(filters?.length ? { filters } : {}) }
}
