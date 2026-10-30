/** An axis's items as shown: the paths of the records let through, items with no data where a field shows them, sorted. */
import type { CellValue, PivotAxisField, PivotCache, PivotFunction, PivotTable } from '../contract.js'
import { sortOrder } from '../sort-order.js'
import type { PivotNode } from './result.js'
import { labelOrder } from './values.js'

/** What the trees are made from: the prepared PivotTable and cache, each record's items, the items let through, the records left, the cube's values. */
export interface Cube {
  table: PivotTable
  cache: PivotCache
  at: Int32Array[]
  shown: (boolean[] | undefined)[]
  alive: readonly number[]
  value(rows: readonly number[], columns: readonly number[], v: number, fn?: PivotFunction): CellValue
  epoch: number
  lists: readonly (readonly string[])[]
}

interface Branch {
  data: boolean
  next: Map<number, Branch>
}

const branch = (data: boolean): Branch => ({ data, next: new Map() })

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

/** The axis's items as nodes, `values` among them where it lies. */
export function axisTree(c: Cube, fields: readonly PivotAxisField[], axis: 'rows' | 'columns'): PivotNode[] {
  const real = fields.filter((f): f is number => f !== 'values')
  const root = branch(true)
  for (const r of c.alive) {
    let b = root
    for (const f of real) {
      const p = c.at[f][r]
      let n = b.next.get(p)
      if (!n) b.next.set(p, (n = branch(true)))
      b = n
    }
  }
  const otherFields = (axis === 'rows' ? c.table.columns : c.table.rows)?.filter((f): f is number => f !== 'values') ?? []
  const nodes = (b: Branch, level: number, path: number[]): PivotNode[] => {
    if (level === real.length) return []
    const f = real[level]
    const pf = c.table.fields[f]
    // items with no data, where the field shows them: every item let through, under every item before
    if (pf.showEmpty)
      (pf.items ?? []).forEach((_, p) => {
        if (!b.next.has(p) && (c.shown[f]?.[p] ?? true)) b.next.set(p, branch(false))
      })
    const items = [...b.next.keys()].sort((x, y) => x - y)
    const sort = pf.autoSort === false ? undefined : pf.sort
    if (sort?.by) {
      const other = otherPath(otherFields, sort.by.at)
      const v = sort.by.value
      const of = (p: number) => (axis === 'rows' ? c.value([...path, p], other, v) : c.value(other, [...path, p], v))
      const val = new Map(items.map((p) => [p, of(p)]))
      items.sort((x, y) => sortOrder(val.get(x), val.get(y), sort.order === 'descending') || x - y)
    } else if (sort) {
      const cached = c.cache.fields[f]?.items ?? []
      const v = (p: number) => {
        const i = pf.items?.[p]?.item
        return i === undefined ? null : (cached[i] ?? null)
      }
      items.sort((x, y) => labelOrder(v(x), v(y), sort.order === 'descending', c.table.customListSort !== false ? c.lists : null, c.epoch) || x - y)
    }
    return items.map((p) => {
      const n = b.next.get(p)!
      const node: PivotNode = { field: f, item: p, children: nodes(n, level + 1, [...path, p]) }
      if (!n.data) node.empty = true
      return node
    })
  }
  const tree = nodes(root, 0, [])
  const at = fields.indexOf('values')
  return at < 0 ? tree : withValues(tree, at, c.table.values?.length ?? 0)
}

const copied = (ns: readonly PivotNode[]): PivotNode[] => ns.map((n) => ({ ...n, children: copied(n.children) }))

/** The tree with a node of each value field at depth `at`, the items below it under each one. */
function withValues(tree: PivotNode[], at: number, count: number): PivotNode[] {
  if (at === 0) return Array.from({ length: count }, (_, v): PivotNode => ({ field: 'values', item: v, children: copied(tree) }))
  return tree.map((n) => ({ ...n, children: withValues(n.children, at - 1, count) }))
}
