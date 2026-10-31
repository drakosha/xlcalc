/**
 * A pivot cache refreshed from its source's cells: fields by the headers, the rows as records, each field's items in
 * the order they first come (text in any case one item, its first spelling kept), and the PivotTables over the cache
 * moved onto the new fields and items, their hidden items kept hidden.
 */
import { caseKey } from '../collate.js'
import type { CellValue, NumberFormat, PivotAxisField, PivotCache, PivotCacheField, PivotField, PivotFieldFacts, PivotTable, PivotValue } from '../contract.js'
import { generalText } from '../numfmt/general.js'
import { ERROR_TEXT } from '../syntax/errors.js'
import { cacheDate, itemKey } from './values.js'

/** A source's cells: the header row first, each cell's value, and whether it shows a date; each column's number format. */
export interface PivotSourceGrid {
  rows: CellValue[][]
  dated?: (row: number, col: number) => boolean
  formats?: (NumberFormat | undefined)[]
}

/** Why a source gives no cache: a header empty (Excel's "The PivotTable field name is not valid"), or no row under the headers. */
export type RefreshRefusal = 'field_name' | 'no_rows'

/** A source's header as a field's name: text as is, a number in General, a boolean or an error by its text; empty for none. */
function headerText(v: CellValue): string {
  if (v === null) return ''
  if (typeof v === 'number') return generalText(v)
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  return typeof v === 'string' ? v : ERROR_TEXT[v.error]
}

/** The headers as fields' names, a repeated one (in any case) numbered from 2 (`Amount`, `Amount2`). */
function namesOf(headers: CellValue[]): string[] | null {
  const taken = new Set<string>()
  const out: string[] = []
  for (const h of headers) {
    const base = headerText(h)
    if (!base) return null
    let name = base
    for (let n = 2; taken.has(caseKey(name)); n++) name = `${base}${n}`
    taken.add(caseKey(name))
    out.push(name)
  }
  return out
}

/** The kinds a field's values hold, with their bounds. */
function factsOf(values: readonly PivotValue[]): PivotFieldFacts {
  const types = new Set<PivotFieldFacts['types'][number]>()
  let integer = true,
    long = false,
    min = Infinity,
    max = -Infinity,
    minDate: string | undefined,
    maxDate: string | undefined
  for (const v of values) {
    if (v === null) types.add('blank')
    else if (typeof v === 'number') {
      types.add('number')
      integer &&= Number.isInteger(v)
      min = Math.min(min, v)
      max = Math.max(max, v)
    } else if (typeof v === 'string') {
      types.add('string')
      long ||= v.length > 255
    } else if (typeof v === 'boolean') types.add('boolean')
    else if ('error' in v) types.add('error')
    else {
      types.add('date')
      if (minDate === undefined || v.date < minDate) minDate = v.date
      if (maxDate === undefined || v.date > maxDate) maxDate = v.date
    }
  }
  const facts: PivotFieldFacts = { types: [...types] }
  if (types.has('number')) {
    if (integer) facts.integer = true
    facts.min = min
    facts.max = max
  } else if (minDate !== undefined) {
    facts.min = { date: minDate }
    facts.max = { date: maxDate! }
  }
  if (long) facts.longText = true
  return facts
}

/** A field is listed (its records index its items) when it holds text, booleans or errors, or a PivotTable groups by it, or it was. */
const listedField = (facts: PivotFieldFacts, grouped: boolean, was: boolean) => grouped || was || facts.types.some((t) => t === 'string' || t === 'boolean' || t === 'error')

/** A cell's value as the cache keeps it: a date where the cell shows one. */
const cacheValue = (v: CellValue, date: boolean, epoch: number): PivotValue => (date && typeof v === 'number' ? { date: cacheDate(v, epoch) } : v)

/** The fields each PivotTable groups by, by index: on its axes or Filters, or with hidden items. */
const groupedBy = (tables: readonly PivotTable[]) =>
  new Set(tables.flatMap((t) => [...(t.rows ?? []), ...(t.columns ?? []), ...(t.pages ?? []).map((p) => p.field), ...t.fields.flatMap((f, i) => (f.items?.some((x) => x.hidden) ? [i] : []))]).values())

/**
 * Refreshes `cache` from `source`, and the PivotTables over it: a field of a header gone leaves them, one of a new header
 * comes on no axis; an item gone from the source is kept as no longer there (`missing`) up to the cache's `retainItems`
 * (all by default, none at 0), new items come last. Fields a grouping or a calculation made are kept after the source's.
 */
export function refreshCache(cache: PivotCache, tables: readonly PivotTable[], source: PivotSourceGrid, epoch = 0): { cache: PivotCache; tables: PivotTable[] } | RefreshRefusal {
  const [headers = [], ...rows] = source.rows
  const names = namesOf(headers)
  if (!names) return 'field_name'
  if (!rows.length) return 'no_rows'
  // the source's fields come first; a field of a grouping's further level or a calculated one ends them
  const made = cache.fields.findIndex((f) => f.group?.base !== undefined || f.formula !== undefined)
  const sourceCount = made < 0 ? cache.fields.length : made
  const old = new Map(cache.fields.slice(0, sourceCount).map((f, i) => [caseKey(f.name), i]))
  const grouped = new Set([...groupedBy(tables)].filter((f): f is number => typeof f === 'number'))
  const fieldMap = new Map<number, number>()
  const itemMaps: (Map<number, number> | undefined)[] = []
  const records: PivotValue[][] = rows.map(() => [])
  const fields = names.map((name, c): PivotCacheField => {
    const was = old.get(caseKey(name))
    if (was !== undefined) fieldMap.set(was, c)
    const prior = was === undefined ? undefined : cache.fields[was]
    const values = rows.map((r, i) => cacheValue(r[c] ?? null, source.dated?.(i + 1, c) ?? false, epoch))
    const facts = factsOf(values)
    const field: PivotCacheField = { ...(prior?.group ? { group: prior.group } : {}), name, facts }
    const format = source.formats?.[c]
    if (format !== undefined) field.numberFormat = format
    if (!listedField(facts, was !== undefined && grouped.has(was), prior?.items !== undefined)) {
      values.forEach((v, i) => (records[i][c] = v))
      return field
    }
    // the items it had, in their order, those still in the source first kept, then the new ones as they come
    const keys = new Map<string, number>()
    const before = prior?.items ?? []
    // each value's first spelling in the source, which an item still there takes (`north` once `North` is gone)
    const present = new Map<string, PivotValue>()
    for (const v of values) if (!present.has(itemKey(v))) present.set(itemKey(v), v)
    const limit = cache.retainItems ?? Infinity
    let retained = 0
    const items: PivotValue[] = []
    const unused: number[] = []
    const map = new Map<number, number>()
    before.forEach((v, i) => {
      const k = itemKey(v)
      if (keys.has(k)) return
      const there = present.has(k)
      if (!there && retained >= limit) return
      if (!there) {
        retained++
        unused.push(items.length)
      }
      keys.set(k, items.length)
      map.set(i, items.length)
      items.push(there ? present.get(k)! : v)
    })
    values.forEach((v, i) => {
      const k = itemKey(v)
      let at = keys.get(k)
      if (at === undefined) {
        keys.set(k, (at = items.length))
        items.push(v)
      }
      records[i][c] = at
    })
    itemMaps[c] = map
    field.items = items
    if (unused.length) field.unused = unused
    return field
  })
  // the fields grouping and calculation made, after the source's
  cache.fields.slice(sourceCount).forEach((_, i) => fieldMap.set(sourceCount + i, fields.length + i))
  for (const f of cache.fields.slice(sourceCount)) {
    const g = f.group
    fields.push(g ? { ...f, group: { ...g, ...(g.base !== undefined ? { base: fieldMap.get(g.base) } : {}), ...(g.parent !== undefined ? { parent: fieldMap.get(g.parent) } : {}) } } : f)
  }
  const { stale: _, ...rest } = cache
  const out: PivotCache = { ...rest, fields, records }
  return { cache: out, tables: tables.map((t) => moved(t, fields, fieldMap, itemMaps)) }
}

/** A PivotTable moved onto the refreshed fields: a field gone leaves every place, items move to their new indexes. */
function moved(t: PivotTable, fields: readonly PivotCacheField[], fieldMap: ReadonlyMap<number, number>, itemMaps: readonly (Map<number, number> | undefined)[]): PivotTable {
  const to = (f: number) => fieldMap.get(f)
  const out: PivotField[] = fields.map(() => ({}))
  // each field's items by their new place on the PivotTable, for the indexes a page field and a sort name
  const placeMaps = new Map<number, Map<number, number>>()
  t.fields.forEach((pf, f) => {
    const n = to(f)
    if (n === undefined) return
    const items = itemMaps[n]
    if (!pf.items || !items) {
      out[n] = pf
      return
    }
    const unused = new Set(fields[n].unused)
    const place = new Map<number, number>()
    const kept: NonNullable<PivotField['items']> = []
    pf.items.forEach((x, p) => {
      if (x.item === undefined) {
        place.set(p, kept.length)
        kept.push(x)
        return
      }
      const i = items.get(x.item)
      if (i === undefined) return
      place.set(p, kept.length)
      const { missing: _, ...rest } = x
      kept.push(unused.has(i) ? { ...rest, item: i, missing: true } : { ...rest, item: i })
    })
    // the items new to the field come last, hidden where the field hides some and does not take new items in
    const listed = new Set(kept.flatMap((x) => (x.item === undefined ? [] : [x.item])))
    const hide = pf.items.some((x) => x.hidden) && !pf.includeNewItems
    fields[n].items!.forEach((_, i) => !listed.has(i) && !unused.has(i) && kept.push(hide ? { item: i, hidden: true } : { item: i }))
    placeMaps.set(n, place)
    out[n] = { ...pf, items: kept }
  })
  const axis = (a?: PivotTable['rows']) => a?.flatMap((f): PivotAxisField[] => (f === 'values' ? [f] : to(f) === undefined ? [] : [to(f)!]))
  const placeOf = (f: number, p: number) => placeMaps.get(f)?.get(p) ?? p
  const result: PivotTable = { ...t, fields: out }
  if (t.rows) result.rows = axis(t.rows)
  if (t.columns) result.columns = axis(t.columns)
  if (t.pages)
    result.pages = t.pages.flatMap((p) => {
      const f = to(p.field)
      if (f === undefined) return []
      let item = p.item === undefined ? undefined : placeMaps.get(f)?.get(p.item)
      // the item shown gone from the source: the first one still there
      if (item !== undefined && out[f].items?.[item]?.missing) item = out[f].items!.findIndex((x) => !x.missing && !x.hidden)
      return [item === undefined || item < 0 ? { field: f } : { field: f, item }]
    })
  if (t.values) result.values = t.values.flatMap((v) => (to(v.field) === undefined ? [] : [{ ...v, field: to(v.field)!, ...(v.baseField !== undefined && to(v.baseField) !== undefined ? { baseField: to(v.baseField) } : {}) }]))
  if (t.filters) result.filters = t.filters.flatMap((x) => (to(x.field) === undefined ? [] : [{ ...x, field: to(x.field)! }]))
  for (const pf of out)
    if (pf.sort?.by?.at) pf.sort = { ...pf.sort, by: { ...pf.sort.by, at: pf.sort.by.at.flatMap((a) => (to(a.field) === undefined ? [] : [{ field: to(a.field)!, item: placeOf(to(a.field)!, a.item) }])) } }
  return result
}
