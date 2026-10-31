/** Excel's label sort: custom-list comparisons and the field-item indexes that follow a new order. */
import { caseKey } from '../collate.js'
import type { PivotArea, PivotCache, PivotTable, PivotValue } from '../contract.js'
import { sortOrder } from '../sort-order.js'
import { groupedItems } from './items.js'
import { sortPivotItems } from './sort.js'
import { errorRank, plainOf } from './values.js'

const indexes = new WeakMap<readonly (readonly string[])[], Map<string, Map<number, number>>>()
function positions(lists: readonly (readonly string[])[]): Map<string, Map<number, number>> {
  let out = indexes.get(lists)
  if (out) return out
  out = new Map()
  lists.forEach((list, i) => list.forEach((text, j) => {
    const key = caseKey(text)
    const at = out!.get(key) ?? new Map<number, number>()
    if (!at.has(i)) at.set(i, j)
    out!.set(key, at)
  }))
  indexes.set(lists, out)
  return out
}

/** Sorts indexes in place; texts belonging to disjoint custom lists compare as less in both directions. */
export function sortLabels(order: number[], value: (i: number) => PivotValue, descending: boolean, lists: readonly (readonly string[])[] | null, epoch: number): void {
  const at = lists && positions(lists)
  const values = new Map(order.map(i => [i, value(i)]))
  const initial = new Map(order.map((i, k) => [i, k]))
  const compare = (x: number, y: number) => {
    const a = values.get(x)!, b = values.get(y)!
    if (at && typeof a === 'string' && typeof b === 'string') {
      const left = at.get(caseKey(a)), right = at.get(caseKey(b))
      if (left) {
        if (right) for (const [list, n] of left) if (right.has(list)) return n - right.get(list)!
        return -1
      }
      if (right) return 1
    }
    return sortOrder(plainOf(a, epoch), plainOf(b, epoch), false) || errorRank(a, b)
  }
  sortPivotItems(order, (pivot, item) => ((compare(item, pivot) || initial.get(item)! - initial.get(pivot)!) < 0) !== descending)
}

/** For each field, each old item index's new index. */
export type ItemOrder = Map<number, number[]>

/** An area over PivotTable item indexes follows those items; cache indexes do not move. */
export function orderedArea(area: PivotArea, maps: ItemOrder): PivotArea {
  if (area.cacheIndex || !area.references) return area
  return { ...area, references: area.references.map(r => {
    const map = typeof r.field === 'number' ? maps.get(r.field) : undefined
    return map && r.items ? { ...r, items: r.items.map(i => map[i] ?? i) } : r
  }) }
}

/** Applies requested label sorts to field items and every table-local reference to their indexes. */
export function orderedLabels(table: PivotTable, cache: PivotCache, requested: boolean | readonly number[], lists: readonly (readonly string[])[], epoch: number): { table: PivotTable; maps: ItemOrder } {
  const maps: ItemOrder = new Map()
  if (requested === false) return { table, maps }
  const fields = table.fields.map((f, i) => {
    if (!f.sort || f.sort.by || !f.items || (requested !== true && !requested.includes(i))) return f
    const values = groupedItems(cache.fields[i]) ?? []
    const order = f.items.map((_, j) => j)
    sortLabels(order, p => values[f.items![p].item!] ?? null, f.sort.order === 'descending', table.customListSort !== false ? lists : null, epoch)
    if (order.every((p, j) => p === j)) return f
    const map: number[] = []
    order.forEach((p, j) => (map[p] = j))
    maps.set(i, map)
    return { ...f, items: order.map(p => f.items![p]) }
  })
  if (!maps.size) return { table, maps }
  const item = (f: number, p: number) => maps.get(f)?.[p] ?? p
  const out: PivotTable = { ...table, fields: fields.map(f => f.sort?.by?.at ? { ...f, sort: { ...f.sort, by: { ...f.sort.by, at: f.sort.by.at.map(a => ({ ...a, item: item(a.field, a.item) })) } } } : f) }
  if (table.pages) out.pages = table.pages.map(p => p.item === undefined ? p : { ...p, item: item(p.field, p.item) })
  if (table.values) out.values = table.values.map(v => v.baseField !== undefined && typeof v.baseItem === 'number' ? { ...v, baseItem: item(v.baseField, v.baseItem) } : v)
  if (table.formats) out.formats = table.formats.map(f => ({ ...f, area: orderedArea(f.area, maps) }))
  delete out.rowLines
  delete out.columnLines
  return { table: out, maps }
}
