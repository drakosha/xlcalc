/**
 * A PivotTable's calculation over its cache (`PivotResult`): the records its filters let through, the value fields
 * summarized at every pair of a path of the rows and a path of the columns, and the axes' items as shown.
 */
import { EPOCH_1904 } from '../dates.js'
import { listsOf } from '../locale/lists.js'
import type { CellValue, ErrorCode, PivotCache, PivotFunction, PivotTable, PivotValue } from '../contract.js'
import { keepsNumbers, Tally } from './aggregate.js'
import { byValue, itemPasses, valuesKept } from './filters.js'
import { axesOf, prepared } from './prepare.js'
import type { PivotComputeOptions, PivotPageResult, PivotResult } from './result.js'
import { axisTree, type Cube } from './tree.js'
import { numberOf } from './values.js'

/** A cell of the cube: the records at it and each value field's tally. */
interface Cell {
  records: number
  tallies: Tally[]
}

/** Each record's item on each grouping field (an index into the PivotTable field's `items`), −1 for a field not grouping. */
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
  const used = new Set<PivotFunction>([...fns, ...table.fields.flatMap((f) => f.subtotals ?? [])])
  const keep = keepsNumbers(used)
  const readers = values.map((v) => valueReader(cache, v.field, epoch))
  const at = itemIndexes(table, cache)
  const shown = shownItems(table, cache, today, epoch)
  const count = cache.records?.length ?? 0

  // the records every field's items let through
  let alive: number[] = []
  for (let r = 0; r < count; r++) if (shown.every((s, f) => !s || (at[f][r] >= 0 && s[at[f][r]]))) alive.push(r)

  const tally = (rs: readonly number[], v: number) => {
    const t = new Tally(keep)
    for (const r of rs) t.add(...readers[v](r))
    return t
  }
  // value filters and Top 10, field by field down the rows then the columns: each item against the others under the same items before it
  for (const fields of [rowFields, colFields]) {
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
          const list = [...items]
          const kept = valuesKept(flt, list.map(([, rs]) => (values[v] ? tally(rs, v).value(fns[v]) : null)))
          list.forEach(([, rs], i) => !kept[i] && rs.forEach((r) => out.add(r)))
        }
        if (out.size) alive = alive.filter((r) => !out.has(r))
      }
    })
  }

  // the cube: each record at every prefix of its row path with every prefix of its column path
  const cube = new Map<string, Cell>()
  for (const r of alive) {
    const rk = prefixes(rowFields.map((f) => at[f][r]))
    const ck = prefixes(colFields.map((f) => at[f][r]))
    for (const a of rk)
      for (const b of ck) {
        const k = `${a}|${b}`
        let cell = cube.get(k)
        if (!cell) cube.set(k, (cell = { records: 0, tallies: values.map(() => new Tally(keep)) }))
        cell.records++
        for (let v = 0; v < values.length; v++) cell.tallies[v].add(...readers[v](r))
      }
  }
  const value = (rows: readonly number[], columns: readonly number[], v: number, fn?: PivotFunction): CellValue =>
    cube.get(`${key(rows)}|${key(columns)}`)?.tallies[v]?.value(fn ?? fns[v]) ?? null
  const hasData = (rows: readonly number[], columns: readonly number[]) => cube.has(`${key(rows)}|${key(columns)}`)
  const view: Cube = { table, cache, at, shown, alive, value, epoch, lists }
  const pages = (table.pages ?? []).map((p): PivotPageResult => {
    const all = (table.fields[p.field].items ?? []).map((_, i) => i)
    const items = p.item !== undefined ? [p.item] : all.filter((i) => !table.fields[p.field].items![i].hidden)
    return { field: p.field, items, all: p.item === undefined && items.length === all.length }
  })
  return {
    table,
    cache,
    rows: { fields: axes.rows, items: axisTree(view, axes.rows, 'rows') },
    columns: { fields: axes.columns, items: axisTree(view, axes.columns, 'columns') },
    pages,
    value,
    hasData,
  }
}

/** The keys of a path's prefixes, the empty one first. */
function prefixes(path: readonly number[]): string[] {
  const out = ['']
  for (let i = 0; i < path.length; i++) out.push(i ? `${out[i]},${path[i]}` : `${path[i]}`)
  return out
}
