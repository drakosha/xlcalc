/**
 * A PivotTable's calculation over its cache (`PivotResult`): the records its filters let through, the value fields
 * summarized at every pair of a path of the rows and a path of the columns, and the axes' items as shown.
 */
import { EPOCH_1904 } from '../dates.js'
import { listsOf } from '../locale/lists.js'
import type { CellValue, ErrorCode, PivotCache, PivotFunction, PivotTable, PivotValue } from '../contract.js'
import { orderedLabels } from './label-sort.js'
import { Tally } from './aggregate.js'
import { Batch, fieldFormulas, itemFormulas, prefetch, settled, sumsOf, withCalculatedItems } from './calculated.js'
import { recordItems } from './group.js'
import { groupedItems } from './items.js'
import { byValue, itemPasses, valuesKept, partitionRecords } from './filters.js'
import { axesOf, prepared } from './prepare.js'
import type { PivotComputeOptions, PivotPageResult, PivotResult } from './result.js'
import { branches, nodesOf, sortBranches, type Branch } from './tree.js'
import { numberOf } from './values.js'

/**
 * Each record's item on each grouping field (an index into the PivotTable field's `items`), −1 for a field not grouping.
 * A grouped field's item is its group's, by its base field's value (`group.ts`).
 */
function itemIndexes(table: PivotTable, cache: PivotCache, epoch: number): Int32Array[] {
  const records = cache.records ?? []
  return table.fields.map((pf, f) => {
    const at = new Int32Array(records.length).fill(-1)
    if (!pf.items || !groupedItems(cache.fields[f])) return at
    const of = new Map<number, number>()
    pf.items.forEach((x, p) => x.item !== undefined && !of.has(x.item) && of.set(x.item, p))
    const item = recordItems(cache, f, epoch)
    records.forEach((r, i) => (at[i] = of.get(item(r)) ?? -1))
    return at
  })
}

/** The items each field lets through: not hidden, the page field's one item, passing the field's label and date filters. */
function shownItems(table: PivotTable, cache: PivotCache, today: number, epoch: number): (boolean[] | undefined)[] {
  const pages = new Map((table.pages ?? []).map((p) => [p.field, p.item]))
  return table.fields.map((pf, f) => {
    const values = groupedItems(cache.fields[f])
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

/** A value as a tally takes a record's. */
const asRecord = (v: CellValue): [n: number | undefined, filled: boolean, error?: ErrorCode] =>
  v === null ? [undefined, false] : typeof v === 'object' ? [undefined, true, v.error] : [typeof v === 'number' ? v : undefined, true]

/**
 * The cache items of field `field` that some record has under the filters of the other fields (hidden items, a page
 * field's item, label and date filters) with a value for a value field, as a slicer of the field tells its items with
 * data (the set slicer: a region whose only record has no price has none under Sum of Price). `table` and `cache` are a
 * calculation's (`PivotResult`), the field listed.
 */
export function itemsWithData(table: PivotTable, cache: PivotCache, field: number, options: PivotComputeOptions = {}): Set<number> {
  const epoch = options.date1904 ? EPOCH_1904 : 0
  const at = itemIndexes(table, cache, epoch)
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
  const ready = prepared(given, givenCache, epoch, lists, options.sortThreads ?? 2)
  const cache = ready.cache
  const ordered = orderedLabels(ready.table, cache, options.sortLabels ?? true, lists, epoch, options.sortThreads ?? 2)
  const table = ordered.table
  const axes = axesOf(table)
  const rowFields = axes.rows.filter((f): f is number => f !== 'values')
  const colFields = axes.columns.filter((f): f is number => f !== 'values')
  const values = table.values ?? []
  const fns = values.map((v) => v.function ?? 'sum')
  const onlySums = fns.every(f => f === 'sum')
  // a calculated field is no record's value: its formula is calculated over the sums of the fields it names at each cell,
  // which are measures after the value fields' (`sums`)
  const formulas = fieldFormulas(cache)
  const sums = sumsOf(values.flatMap((v) => (formulas.has(v.field) ? [v.field] : [])), formulas)
  const sumAt = new Map(sums.map((f, i) => [f, values.length + i]))
  const measures = [...values.map((v) => v.field), ...sums]
  const readers = measures.map((f) => valueReader(cache, f, epoch))
  const at = itemIndexes(table, cache, epoch)
  const shown = shownItems(table, cache, today, epoch)
  const count = cache.records?.length ?? 0

  // the records every field's items let through
  let alive = partitionRecords(count, shown.some(Boolean) ? r => shown.every((s, f) => !s || (at[f][r] >= 0 && s[at[f][r]])) : undefined)

  let recordOrder: Int32Array | undefined
  const cacheItem = (f: number, p: number) => table.fields[f].items?.[p]?.item ?? p
  const addRecord = (t: Tally, m: number, r: number) => {
    const [n, filled, error] = readers[m](r)
    if (error && !recordOrder) {
      recordOrder = new Int32Array(count)
      alive.forEach((id, i) => (recordOrder![id] = i))
    }
    t.add(n, filled, error, error ? [...rowFields.map(f => cacheItem(f, at[f][r])), ...colFields.map(f => cacheItem(f, at[f][r])), recordOrder![r]] : undefined)
  }
  const tally = (rs: readonly number[], v: number) => {
    const t = new Tally()
    for (const r of rs) addRecord(t, v, r)
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
  const leafRecords = new Map<string, number[]>()
  for (const r of alive) {
    const k = `${key(rowFields.map((f) => at[f][r]))}|${key(colFields.map((f) => at[f][r]))}`
    ;(leafRecords.get(k) ?? leafRecords.set(k, []).get(k)!).push(r)
    let cell = leaves.get(k)
    if (!cell) leaves.set(k, (cell = measures.map((_, m) => new Tally(m >= values.length))))
    for (let m = 0; m < measures.length; m++) addRecord(cell[m], m, r)
  }
  const sorting = { table, cache, shown, epoch, lists, sortThreads: options.sortThreads ?? 2 }
  const rowRoot = branches(sorting, at, shownBy.rows, rowFields)
  const colRoot = branches(sorting, at, shownBy.columns, colFields)
  const calcItems = itemFormulas(table, cache, epoch)
  withCalculatedItems(rowRoot, rowFields, calcItems, shown)
  withCalculatedItems(colRoot, colFields, calcItems, shown)
  const batch = new Batch()
  let memo = new Map<string, Tally[] | null>()
  let pathsMemo = new WeakMap<Branch, number[][]>()
  const lastPaths = (b: Branch, path: readonly number[], fields: readonly number[]): number[][] => {
    let paths = pathsMemo.get(b)
    if (!paths) {
      paths = path.length === fields.length ? [[...path]] : b.order.flatMap(p => lastPaths(b.next.get(p)!, [...path, p], fields))
      pathsMemo.set(b, paths)
    }
    return paths
  }
  const newTallies = () => measures.map((_, m) => new Tally(m >= values.length))
  const walk = (root: Branch, path: readonly number[]) => path.reduce<Branch | undefined>((b, p) => b?.next.get(p), root)
  const cell = (rows: readonly number[], rb: Branch, cols: readonly number[], cb: Branch, sumMode: boolean): Tally[] | null => {
    const plain = `${key(rows)}|${key(cols)}`
    const k = `${plain}|${sumMode ? 1 : 0}`
    const cached = memo.get(k)
    if (cached !== undefined) return cached
    let out: Tally[] | null
    if (rows.length === rowFields.length && cols.length === colFields.length) out = leaves.get(plain) ?? null
    else {
      const rp = lastPaths(rb, rows, rowFields), cp = lastPaths(cb, cols, colFields)
      // Sum combines leaf-cell totals across each finest row, then the rows. Other measures replay the records;
      // a total over columns first accumulates each finest row, then combines those row states.
      const byRow = sumMode || cols.length < colFields.length
      out = null
      for (const r of rp) {
        let row: Tally[] | null = byRow ? null : out
        for (const c of cp) {
          const leaf = `${key(r)}|${key(c)}`
          if (sumMode) {
            let t = leaves.get(leaf)
            if (!t && calcAt(r, c)) {
              t = newTallies()
              t.forEach((x, m) => x.add(...asRecord(measure(r, c, m, 'sum'))))
            }
            if (t) { row ??= newTallies(); row.forEach((x, m) => x.merge(t[m])) }
          } else {
            if (!leafRecords.has(leaf) && calcAt(r, c)) {
              row ??= newTallies()
              row.forEach((x, m) => x.add(...asRecord(measure(r, c, m, m < values.length ? fns[m] : 'sum'))))
            }
            for (const record of leafRecords.get(leaf) ?? []) {
              row ??= newTallies()
              for (let m = 0; m < measures.length; m++) addRecord(row[m], m, record)
            }
          }
        }
        if (row) {
          if (byRow) { out ??= newTallies(); out.forEach((x, m) => x.merge(row![m])) }
          else out = row
        }
      }
    }
    memo.set(k, out)
    return out
  }
  const reach = (rows: readonly number[], columns: readonly number[], sumMode = false) => {
    const [rb, cb] = [walk(rowRoot, rows), walk(colRoot, columns)]
    return rb?.data && cb?.data ? cell(rows, rb, columns, cb, sumMode) : null
  }
  // the first calculated item of a path: its axis, level and formula
  const calcAt = (rows: readonly number[], columns: readonly number[]) => {
    for (const [path, fields] of [[rows, rowFields], [columns, colFields]] as const)
      for (let k = 0; k < path.length; k++) {
        const t = calcItems.get(fields[k])?.get(path[k])
        if (t) return { rows: path === rows, k, t }
      }
    return null
  }
  const reached = (rows: readonly number[], columns: readonly number[]) => !!(walk(rowRoot, rows)?.data && walk(colRoot, columns)?.data)
  /** A measure's value at the paths: a value field's by `fn`, a calculated field's formula, a calculated item's. */
  const measure = (rows: readonly number[], columns: readonly number[], m: number, fn: PivotFunction): CellValue => {
    const field = measures[m]
    if (formulas.has(field)) return fieldValue(rows, columns, field)
    const c = calcAt(rows, columns)
    if (!c) return reach(rows, columns, fn === 'sum' && m < values.length && onlySums)?.[m]?.value(fn) ?? null
    if (!reached(rows, columns)) return null
    const swap = (q: number) => (c.rows ? [rows.map((p, k) => (k === c.k ? q : p)), columns] : [rows, columns.map((p, k) => (k === c.k ? q : p))]) as [number[], number[]]
    return batch.get(c.t.text, c.t.args.map((q) => measure(...swap(q), m, fn)))
  }
  /** A field's value at a cell: its sum, or a calculated field's formula over the values of the fields it names. */
  const fieldValue = (rows: readonly number[], columns: readonly number[], field: number): CellValue => {
    const t = formulas.get(field)
    // a field with no number at the cell is 0 (the set calc-held: MIN(Price,10*Qty) over no price is 0)
    if (!t) return measure(rows, columns, sumAt.get(field)!, 'sum') ?? 0
    // every cell of the rows and columns shown, one no record reaches too (its sums blank)
    if (!reached(rows, columns)) return null
    return batch.get(t.text, t.args.map((f) => fieldValue(rows, columns, f)))
  }
  // a calculated field's subtotals by functions of their own are empty (the set calc)
  const value = (rows: readonly number[], columns: readonly number[], v: number, fn?: PivotFunction): CellValue =>
    fn !== undefined && formulas.has(measures[v]) ? null : settled(batch, () => measure(rows, columns, v, fn ?? fns[v]))
  const hasData = (rows: readonly number[], columns: readonly number[]) => (calcAt(rows, columns) ? reached(rows, columns) : reach(rows, columns) !== null)
  sortBranches(sorting, rowRoot, rowFields, colFields, true, value)
  sortBranches(sorting, colRoot, colFields, rowFields, false, value)
  memo = new Map()
  pathsMemo = new WeakMap()
  if (formulas.size || calcItems.size) prefetch(batch, rowRoot, colRoot, values.length, (r, c, v) => measure(r, c, v, fns[v]))
  const pages = (table.pages ?? []).map((p): PivotPageResult => {
    const all = (table.fields[p.field].items ?? []).map((_, i) => i)
    const items = p.item !== undefined ? [p.item] : all.filter((i) => !table.fields[p.field].items![i].hidden)
    return { field: p.field, items, all: p.item === undefined && items.length === all.length }
  })
  return {
    table,
    itemOrder: ordered.maps,
    cache,
    rows: { fields: axes.rows, items: nodesOf(rowRoot, axes.rows, values.length) },
    columns: { fields: axes.columns, items: nodesOf(colRoot, axes.columns, values.length) },
    pages,
    value,
    hasData,
  }
}

