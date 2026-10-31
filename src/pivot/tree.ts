/**
 * An axis's items as shown: the paths of the records let through as branches, items with no data where a field shows
 * them, each branch's items in the order shown (sorted by label or by a value), then as `PivotNode`s.
 */
import type { CellValue, ErrorCode, PivotAxisField, PivotCache, PivotTable } from '../contract.js'
import { ERROR_TYPE } from '../error-type.js'
import { sortOrder } from '../sort-order.js'
import type { PivotNode } from './result.js'
import { groupedItems } from './group.js'
import { labelOrder } from './values.js'

/** An item of an axis and the items of the next field under it, in the order shown (`order`); `data`: some record reaches it. */
export interface Branch {
  data: boolean
  next: Map<number, Branch>
  order: number[]
}

const branch = (data: boolean): Branch => ({ data, next: new Map(), order: [] })

/** What sorting reads: the prepared PivotTable and cache, the items let through, the epoch and custom lists. */
export interface Sorting {
  table: PivotTable
  cache: PivotCache
  shown: (boolean[] | undefined)[]
  epoch: number
  lists: readonly (readonly string[])[]
}

/** The branches of an axis's fields `real` from each record's items, with the items of no data where a field shows them, in the PivotTable's order. */
export function branches(s: Sorting, at: readonly Int32Array[], alive: readonly number[], real: readonly number[]): Branch {
  const root = branch(true)
  for (const r of alive) {
    let b = root
    for (const f of real) {
      const p = at[f][r]
      let n = b.next.get(p)
      if (!n) b.next.set(p, (n = branch(true)))
      b = n
    }
  }
  const fill = (b: Branch, level: number) => {
    if (level === real.length) return
    const f = real[level]
    // items with no data, where the field shows them: every item let through, under every item before
    if (s.table.fields[f].showEmpty)
      (s.table.fields[f].items ?? []).forEach((_, p) => {
        if (!b.next.has(p) && (s.shown[f]?.[p] ?? true)) b.next.set(p, branch(false))
      })
    b.order = [...b.next.keys()].sort((x, y) => x - y)
    for (const n of b.next.values()) fill(n, level + 1)
  }
  fill(root, 0)
  return root
}

/**
 * The order of items by a value: an empty value the least, numbers, then errors last either way, the later of
 * ERROR.TYPE first going up, the earlier going down (the set sort: `#N/A` before `#DIV/0!` ascending).
 */
function valueOrder(a: CellValue, b: CellValue, desc: boolean): number {
  const ea = a !== null && typeof a === 'object'
  const eb = b !== null && typeof b === 'object'
  if (ea || eb) return ea && eb ? (desc ? 1 : -1) * ((ERROR_TYPE[(a as { error: ErrorCode }).error] ?? 0) - (ERROR_TYPE[(b as { error: ErrorCode }).error] ?? 0)) : ea ? 1 : -1
  const c = a === b ? 0 : a === null ? -1 : b === null ? 1 : sortOrder(a, b, false)
  return desc ? -c : c
}

/** The path on the other axis that a sort by value reads its values at: its fields' items given, up to the first not given. */
function otherPath(fields: readonly number[], given: readonly { field: number; item: number }[] = []): number[] {
  const out: number[] = []
  for (const f of fields) {
    const g = given.find((x) => x.field === f)
    if (!g) break
    out.push(g.item)
  }
  return out
}

/**
 * Sorts each branch's items by its field's sort: by label (custom lists first, with `customListSort`), or by a value
 * field's values at the other axis's items (`value`, its paths on the rows and the columns); ties under an item of a field
 * before go by the item's value over all of them (`alone`), then keep the order before.
 */
export function sortBranches(s: Sorting, root: Branch, real: readonly number[], other: readonly number[], rows: boolean, value: (path: number[], at: number[], v: number) => CellValue, alone: (f: number, item: number, at: number[], v: number) => CellValue): void {
  const walk = (b: Branch, level: number, path: number[]) => {
    if (level === real.length) return
    const f = real[level]
    const pf = s.table.fields[f]
    const sort = pf.sort
    if (sort?.by) {
      const at = otherPath(other, sort.by.at)
      const v = sort.by.value
      const desc = sort.order === 'descending'
      const val = new Map(b.order.map((p) => [p, rows ? value([...path, p], at, v) : value(at, [...path, p], v)]))
      // under an item of a field before, ties go by the items' values over every item before (the set sort-held)
      const whole = new Map(level ? b.order.map((p) => [p, alone(f, p, at, v)]) : [])
      const was = new Map(b.order.map((p, i) => [p, i]))
      b.order.sort((x, y) => valueOrder(val.get(x)!, val.get(y)!, desc) || (level ? valueOrder(whole.get(x)!, whole.get(y)!, desc) : 0) || was.get(x)! - was.get(y)!)
    } else if (sort) {
      const cached = groupedItems(s.cache.fields[f]) ?? []
      const item = (p: number) => {
        const i = pf.items?.[p]?.item
        return i === undefined ? null : (cached[i] ?? null)
      }
      b.order.sort((x, y) => labelOrder(item(x), item(y), sort.order === 'descending', s.table.customListSort !== false ? s.lists : null, s.epoch) || x - y)
    }
    for (const p of b.order) walk(b.next.get(p)!, level + 1, [...path, p])
  }
  walk(root, 0, [])
}

/** The axis's items as nodes, `values` among them where it lies. */
export function nodesOf(root: Branch, fields: readonly PivotAxisField[], values: number): PivotNode[] {
  const real = fields.filter((f): f is number => f !== 'values')
  const nodes = (b: Branch, level: number): PivotNode[] =>
    level === real.length
      ? []
      : b.order.map((p) => {
          const n = b.next.get(p)!
          const node: PivotNode = { field: real[level], item: p, children: nodes(n, level + 1) }
          if (!n.data) node.empty = true
          return node
        })
  const tree = nodes(root, 0)
  const at = fields.indexOf('values')
  return at < 0 ? tree : withValues(tree, at, values)
}

const copied = (ns: readonly PivotNode[]): PivotNode[] => ns.map((n) => ({ ...n, children: copied(n.children) }))

/** The tree with a node of each value field at depth `at`, the items below it under each one. */
function withValues(tree: PivotNode[], at: number, count: number): PivotNode[] {
  if (at === 0) return Array.from({ length: count }, (_, v): PivotNode => ({ field: 'values', item: v, children: copied(tree) }))
  return tree.map((n) => ({ ...n, children: withValues(n.children, at - 1, count) }))
}
