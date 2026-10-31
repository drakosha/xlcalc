/**
 * A slicer's items as Excel shows them after its PivotTables are calculated: chosen unless hidden on the first
 * PivotTable it filters, marked when no record has them under the other fields' filters, in its order (by label, the
 * custom lists first, ascending by default) with the items with data first unless its cache says otherwise.
 */
import type { PivotCache, PivotTable, SlicerCache } from '../contract.js'
import { fieldIndex } from '../structure/slicer-choice.js'
import { itemsWithData } from './compute.js'
import type { PivotComputeOptions } from './result.js'
import { labelOrder } from './values.js'
import { EPOCH_1904 } from '../dates.js'

/** A slicer cache's items from the calculations of its PivotTables (`results`, the first the one it reads its choice from); the cache as it was without one. */
export function slicerItems(sc: SlicerCache, results: readonly { table: PivotTable; cache: PivotCache }[], options: PivotComputeOptions): SlicerCache {
  const first = results[0]
  const f = first ? fieldIndex(first.cache, sc.field) : -1
  const values = f >= 0 ? first.cache.fields[f].items : undefined
  if (!values) return sc
  const data = new Set<number>()
  for (const r of results) for (const k of itemsWithData(r.table, r.cache, f, options)) data.add(k)
  const hidden = new Set((first.table.fields[f]?.items ?? []).flatMap((x) => (x.hidden && x.item !== undefined ? [x.item] : [])))
  const epoch = options.date1904 ? EPOCH_1904 : 0
  const lists = sc.customListSort !== false ? (options.lists ?? null) : null
  const order = values.map((_, k) => k).sort((a, b) => labelOrder(values[a], values[b], sc.sort === 'descending', lists, epoch) || a - b)
  // Excel's default: the items with data at the top
  if (sc.noData?.mark !== false && sc.noData?.last !== false) order.sort((a, b) => Number(!data.has(a)) - Number(!data.has(b)))
  return { ...sc, items: order.map((item) => ({ item, ...(hidden.has(item) ? { selected: false } : {}), ...(data.has(item) ? {} : { noData: true }) })) }
}
