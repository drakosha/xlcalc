/**
 * A PivotTable's calculation over its cache (`PivotResult`): the records its filters let through, the value fields
 * summarized at every pair of a path of the rows and a path of the columns, and the axes' items as shown.
 */
import { EPOCH_1904 } from '../dates.js'
import { listsOf } from '../locale/lists.js'
import type { CellValue, ErrorCode, PivotCache, PivotFunction, PivotTable, PivotValue } from '../contract.js'
import { Tally } from './aggregate.js'
import { byValue, itemPasses, valuesKept } from './filters.js'
import { axesOf, prepared } from './prepare.js'
import type { PivotComputeOptions, PivotPageResult, PivotResult } from './result.js'
import { branches, nodesOf, sortBranches, type Branch } from './tree.js'
import { numberOf } from './values.js'

/**
 * Each record's item on each grouping field (an index into the PivotTable field's `items`), −1 for a field not grouping.
 * A field a grouping made (`PivotCacheField.group`) has no value in the records: its item is to come from its base field's.
 */
function itemIndexes(table: PivotTable, cache: PivotCache): Int32Array[] {
  const records = cache.records ?? []
  return table.fields.map((pf, f) => {
    const at = new Int32Array(records.length).fill(-1)
    if (!pf.items || !cache.fields[f]?.items) return at
    const of = new Map<number, number>()
    pf.items.forEach((x, p) => x.item !== undefined && !of.has(x.item) && of.set(x.item, p))
    records.forEach((r, i) => (at[i] = of.get(r[f] as number) ?? -1))
    return at
  })
}

/** The items each field lets through: not hidden, the page field's one item, passing the field's label and date filters. */
function shownItems(table: PivotTable, cache: PivotCache, today: number, epoch: number): (boolean[] | undefined)[] {
  const pages = new Map((table.pages ?? []).map((p) => [p.field, p.item]))
  return table.fields.map((pf, f) => {
    const values = cache.fields[f]?.items
    if (!pf.items || !values) return undefined
    const page = pages.get(f)
    const filters = (table.filters ?? []).filter((x) => x.field === f && !byValue(x))
    const shown = pf.items.map((x, p) => {
      if (page !== undefined ? p !== page : x.hidden) return false
      const v = x.item === undefined ? null : values[x.item]
      return filters.every((flt) => itemPasses(flt, v, x.caption, today, epoch))
    })
    return shown.every(Boolean) ? undefined : shown
  })
}

/** The value of a record's field as a tally takes it. */
function valueReader(cache: PivotCache, field: number, epoch: number): (r: number) => [n: number | undefined, filled: boolean, error?: ErrorCode] {
  const records = cache.records ?? []
  const items = cache.fields[field]?.items
  return (r) => {
    const raw = records[r]?.[field] ?? null
    const v: PivotValue = items && typeof raw === 'number' ? (items[raw] ?? null) : raw
    if (v === null) return [undefined, false]
    if (typeof v === 'object' && 'error' in v) return [undefined, true, v.error]
    return [numberOf(v, epoch), true]
  }
}

const key = (path: readonly number[]) => path.join(',')

/**
 * The cache items of field `field` that some record has under the filters of the other fields (hidden items, a page
 * field's item, label and date filters) with a value for a value field, as a slicer of the field tells its items with
 * data (the set slicer: a region whose only record has no price has none under Sum of Price). `table` and `cache` are a
 * calculation's (`PivotResult`), the field listed.
 */
export function itemsWithData(table: PivotTable, cache: PivotCache, field: number, options: PivotComputeOptions = {}): Set<number> {
  const epoch = options.date1904 ? EPOCH_1904 : 0
  const at = itemIndexes(table, cache)
  const shown = shownItems(table, cache, options.today ?? 0, epoch)
  const readers = (table.values ?? []).map((v) => valueReader(cache, v.field, epoch))
  const out = new Set<number>()
  ;(cache.records ?? []).forEach((rec, r) => {
    if (typeof rec[field] !== 'number' || out.has(rec[field] as number)) return
    if (shown.every((s, f) => f === field || !s || (at[f][r] >= 0 && s[at[f][r]])) && (!readers.length || readers.some((read) => read(r)[1]))) out.add(rec[field] as number)
  })
  return out
}

/** Calculates a PivotTable over its cache's records. */
export function computePivot(given: PivotTable, givenCache: PivotCache, options: PivotComputeOptions = {}): PivotResult {
  const epoch = options.date1904 ? EPOCH_1904 : 0
  const today = options.today ?? 0
  const lists = options.lists ?? listsOf({})
  const { table, cache } = prepared(given, givenCache, epoch, lists)
  const axes = axesOf(table)
  const rowFields = axes.rows.filter((f): f is number => f !== 'values')
  const colFields = axes.columns.filter((f): f is number => f !== 'values')
  const values = table.values ?? []
  const fns = values.map((v) => v.function ?? 'sum')
  // a calculated field is no record's value: its formula is to be calculated over the sums of the fields it names at each cell
  const readers = values.map((v) => valueReader(cache, v.field, epoch))
  const at = itemIndexes(table, cache)
  const shown = shownItems(table, cache, today, epoch)
  const count = cache.records?.length ?? 0

  // the records every field's items let through
  let alive: number[] = []
  for (let r = 0; r < count; r++) if (shown.every((s, f) => !s || (at[f][r] >= 0 && s[at[f][r]]))) alive.push(r)

  const tally = (rs: readonly number[], v: number) => {
    const t = new Tally()
    for (const r of rs) t.add(...readers[v](r))
    return t
  }
  // value filters and Top 10, field by field down the rows then the columns: each item against the others under the same
  // items before it; an axis shows the items its own filters leave, the values are of the records both leave (the set
  // filter-held: rows all filtered out by value, the columns' items still shown)
  const shownBy = { rows: alive, columns: alive }
  for (const [fields, axis] of [[rowFields, 'rows'], [colFields, 'columns']] as const) {
    fields.forEach((f, level) => {
      for (const flt of (table.filters ?? []).filter((x) => x.field === f && byValue(x))) {
        const groups = new Map<string, Map<number, number[]>>()
        for (const r of alive) {
          const parent = key(fields.slice(0, level).map((g) => at[g][r]))
          const items = groups.get(parent) ?? groups.set(parent, new Map()).get(parent)!
          const item = at[f][r]
          ;(items.get(item) ?? items.set(item, []).get(item)!).push(r)
        }
        const v = flt.value ?? 0
        const out = new Set<number>()
        for (const items of groups.values()) {
          const list = [...items].sort(([x], [y]) => x - y)
          const kept = valuesKept(flt, list.map(([, rs]) => (values[v] ? tally(rs, v).value(fns[v]) : null)))
          list.forEach(([, rs], i) => !kept[i] && rs.forEach((r) => out.add(r)))
        }
        if (!out.size) continue
        alive = alive.filter((r) => !out.has(r))
        shownBy[axis] = shownBy[axis].filter((r) => !out.has(r))
      }
    })
  }

  // the cells of the cube's last items on both axes, of the records in their order
  const leaves = new Map<string, Tally[]>()
  for (const r of alive) {
    const k = `${key(rowFields.map((f) => at[f][r]))}|${key(colFields.map((f) => at[f][r]))}`
    let cell = leaves.get(k)
    if (!cell) leaves.set(k, (cell = values.map(() => new Tally())))
    for (let v = 0; v < values.length; v++) cell[v].add(...readers[v](r))
  }
  const sorting = { table, cache, shown, epoch, lists }
  const rowRoot = branches(sorting, at, shownBy.rows, rowFields)
  const colRoot = branches(sorting, at, shownBy.columns, colFields)
  // a total is made of the cells under it in the cache's order of their items, down the rows first, then across the
  // columns, as Excel adds them (the set basic: a product's grand total)
  const cacheOrder = (f: number, b: Branch) => [...b.next.keys()].sort((x, y) => (table.fields[f].items?.[x]?.item ?? x) - (table.fields[f].items?.[y]?.item ?? y))
  let memo = new Map<string, Tally[] | null>()
  const walk = (root: Branch, path: readonly number[]) => path.reduce<Branch | undefined>((b, p) => b?.next.get(p), root)
  const cell = (rows: readonly number[], rb: Branch, cols: readonly number[], cb: Branch): Tally[] | null => {
    const k = `${key(rows)}|${key(cols)}`
    let out = memo.get(k)
    if (out !== undefined) return out
    if (rows.length === rowFields.length && cols.length === colFields.length) out = leaves.get(k) ?? null
    else {
      const down = rows.length < rowFields.length
      const b = down ? rb : cb
      out = null
      for (const p of cacheOrder(down ? rowFields[rows.length] : colFields[cols.length], b)) {
        const n = b.next.get(p)!
        if (!n.data) continue
        const t = down ? cell([...rows, p], n, cols, cb) : cell(rows, rb, [...cols, p], n)
        if (!t) continue
        out ??= values.map(() => new Tally())
        out.forEach((x, v) => x.merge(t[v]))
      }
    }
    memo.set(k, out)
    return out
  }
  const reach = (rows: readonly number[], columns: readonly number[]) => {
    const [rb, cb] = [walk(rowRoot, rows), walk(colRoot, columns)]
    return rb?.data && cb?.data ? cell(rows, rb, columns, cb) : null
  }
  const value = (rows: readonly number[], columns: readonly number[], v: number, fn?: PivotFunction): CellValue => reach(rows, columns)?.[v]?.value(fn ?? fns[v]) ?? null
  const hasData = (rows: readonly number[], columns: readonly number[]) => reach(rows, columns) !== null
  // sorting by values reads them in the order before; the cells are made again in the order shown
  // an item's value over every item of the fields before it on its axis, at the other axis's path `other`
  const alone = (f: number, item: number, other: number[], v: number, otherFields: readonly number[]): CellValue =>
    values[v] ? tally(alive.filter((r) => at[f][r] === item && other.every((p, k) => at[otherFields[k]][r] === p)), v).value(fns[v]) : null
  sortBranches(sorting, rowRoot, rowFields, colFields, true, value, (f, p, o, v) => alone(f, p, o, v, colFields))
  sortBranches(sorting, colRoot, colFields, rowFields, false, value, (f, p, o, v) => alone(f, p, o, v, rowFields))
  memo = new Map()
  const pages = (table.pages ?? []).map((p): PivotPageResult => {
    const all = (table.fields[p.field].items ?? []).map((_, i) => i)
    const items = p.item !== undefined ? [p.item] : all.filter((i) => !table.fields[p.field].items![i].hidden)
    return { field: p.field, items, all: p.item === undefined && items.length === all.length }
  })
  return {
    table,
    cache,
    rows: { fields: axes.rows, items: nodesOf(rowRoot, axes.rows, values.length) },
    columns: { fields: axes.columns, items: nodesOf(colRoot, axes.columns, values.length) },
    pages,
    value,
    hasData,
  }
}

