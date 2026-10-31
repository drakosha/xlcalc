/**
 * A cache's fields added and removed with the PivotTables over it following: a field added comes last with an empty
 * field on each PivotTable; a field removed leaves every place that names it, and the fields after it move down.
 */
import type { PivotArea, PivotAxisField, PivotCache, PivotCacheField, PivotTable } from '../contract.js'

/** A cache and the PivotTables over it, which an operation on the cache changes together. */
export interface CacheTables {
  cache: PivotCache
  tables: PivotTable[]
}

/** The fields `added` at the end of the cache, an empty field each on every PivotTable; their first index. */
export function withFields(x: CacheTables, added: PivotCacheField[]): number {
  const at = x.cache.fields.length
  x.cache = { ...x.cache, fields: [...x.cache.fields, ...added] }
  x.tables = x.tables.map((t) => ({ ...t, fields: [...t.fields, ...added.map(() => ({}))] }))
  return at
}

/** The fields `removed` out of the cache and of every place the PivotTables name them. */
export function withoutFields(x: CacheTables, removed: ReadonlySet<number>): void {
  if (!removed.size) return
  const map = new Map<number, number>()
  let n = 0
  x.cache.fields.forEach((_, i) => !removed.has(i) && map.set(i, n++))
  const to = (f: number) => map.get(f)
  const fields = x.cache.fields.flatMap((f, i) => {
    if (removed.has(i)) return []
    const g = f.group
    if (!g) return [f]
    const group = { ...g }
    for (const k of ['base', 'parent'] as const) if (group[k] !== undefined) (to(group[k]!) === undefined ? delete group[k] : (group[k] = to(group[k]!)))
    // a field whose only grouping was a parent gone is no longer grouped
    return [Object.keys(group).length ? { ...f, group } : (({ group: _, ...rest }) => rest)(f)]
  })
  const calculatedItems = x.cache.calculatedItems?.flatMap((c) => (to(c.field) === undefined ? [] : [{ ...c, field: to(c.field)! }]))
  x.cache = { ...x.cache, fields, ...(calculatedItems ? { calculatedItems } : {}) }
  x.tables = x.tables.map((t) => movedTable(t, to))
}

const axis = (a: PivotAxisField[] | undefined, to: (f: number) => number | undefined) => a?.flatMap((f): PivotAxisField[] => (f === 'values' ? [f] : to(f) === undefined ? [] : [to(f)!]))

function movedArea(a: PivotArea, to: (f: number) => number | undefined): PivotArea | null {
  if (typeof a.field === 'number' && to(a.field) === undefined) return null
  const out: PivotArea = { ...a }
  if (typeof a.field === 'number') out.field = to(a.field)
  if (a.references) out.references = a.references.flatMap((r) => (r.field === 'values' ? [r] : to(r.field) === undefined ? [] : [{ ...r, field: to(r.field)! }]))
  return out
}

/** A PivotTable with its fields moved by `to` (undefined for a field gone). */
function movedTable(t: PivotTable, to: (f: number) => number | undefined): PivotTable {
  const out: PivotTable = { ...t, fields: t.fields.filter((_, i) => to(i) !== undefined) }
  out.fields = out.fields.map((pf) => (pf.sort?.by?.at ? { ...pf, sort: { ...pf.sort, by: { ...pf.sort.by, at: pf.sort.by.at.flatMap((a) => (to(a.field) === undefined ? [] : [{ ...a, field: to(a.field)! }])) } } } : pf))
  if (t.rows) out.rows = axis(t.rows, to)
  if (t.columns) out.columns = axis(t.columns, to)
  if (t.pages) out.pages = t.pages.flatMap((p) => (to(p.field) === undefined ? [] : [{ ...p, field: to(p.field)! }]))
  if (t.values) out.values = t.values.flatMap((v) => (to(v.field) === undefined ? [] : [{ ...v, field: to(v.field)!, ...(v.baseField !== undefined ? { baseField: to(v.baseField) ?? 0 } : {}) }]))
  if (t.filters) out.filters = t.filters.flatMap((f) => (to(f.field) === undefined ? [] : [{ ...f, field: to(f.field)! }]))
  if (t.formats) out.formats = t.formats.flatMap((f) => {
    const area = movedArea(f.area, to)
    return area ? [{ ...f, area }] : []
  })
  return out
}

/** The axis (`rows` or `columns`) a field lies on in a PivotTable, and its place there. */
export function placeOf(t: PivotTable, f: number): { axis: 'rows' | 'columns'; at: number } | null {
  for (const axis of ['rows', 'columns'] as const) {
    const at = t[axis]?.indexOf(f) ?? -1
    if (at >= 0) return { axis, at }
  }
  return null
}

/** The PivotTable with `fields` put on an axis at `at`, in place of `count` fields there. */
export function placed(t: PivotTable, where: { axis: 'rows' | 'columns'; at: number }, fields: number[], count = 0): PivotTable {
  const list = [...(t[where.axis] ?? [])]
  list.splice(where.at, count, ...fields)
  return { ...t, [where.axis]: list }
}
