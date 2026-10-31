/**
 * Show Values As: a value field's cell as a share of a total or of its parent, set against a base item (a share, a
 * difference), a running total over the base field's items, a rank among them, or an index, calculated over the
 * values of the cube (`PivotResult.value`, which answers for any items). The layout takes its values from here.
 *
 * Excel's rules (the set `show`): a cell with no record counts as 0, so it shows 0% of a total, a running total so far,
 * a difference from its base; a share of a base item or a difference in percent shows #NULL! there. A share of an
 * empty cell (a parent, a base item, the running total of empty cells) is empty, of 0 or of an empty total of the row
 * or column #DIV/0!. A base item missing from the line (no line of the base item with the same items under it) is
 * #N/A; the previous or next item is the previous or next that has such a line, the item itself for the first or
 * last; a base item not shown (hidden, filtered out) and a base field on no axis are #N/A in every cell. The cells of
 * the base item itself are empty for a difference, 100% for a share; the lines without the base field (totals above
 * it) are empty. A running total takes the last error it passes; an error has no rank, and ties share one, the next
 * value taking the next rank; a share of a parent is empty where the value or its parent is an error. Hidden and
 * filtered items take no part: totals, running totals, ranks and the previous and next items are those of the items
 * shown (the sets show-items and show-held).
 */
import type { CellValue, ErrorCode, PivotDataField, PivotFunction } from '../contract.js'
import { pivotAdd, pivotDivide, pivotMultiply } from './number.js'
import type { PivotNode, PivotResult } from './result.js'

/** A value at the paths of the rows and the columns, of a value field, by a subtotal's own function. */
export type PivotValueAt = (rows: readonly number[], columns: readonly number[], value: number, fn?: PivotFunction) => CellValue

const NA: CellValue = { error: 'NA' }
const NULL: CellValue = { error: 'NULL' }
const DIV0: CellValue = { error: 'DIV0' }

const isError = (x: CellValue): x is { error: ErrorCode } => x !== null && typeof x === 'object'

/** `x / y`: an error of either, empty over an empty `y`, #DIV/0! over 0. */
function ratio(x: CellValue, y: CellValue): CellValue {
  if (isError(x)) return x
  if (isError(y)) return y
  if (typeof x !== 'number' || typeof y !== 'number') return null
  return y === 0 ? DIV0 : pivotDivide(x, y)
}

/** A share of a parent: empty where either is an error (the set show-held). */
const ofParent = (x: CellValue, parent: CellValue): CellValue => (isError(x) || isError(parent) ? null : ratio(x, parent))

/** The items under a node's path, `values` passed over (its copies under each value field are alike). */
function under(nodes: readonly PivotNode[]): readonly PivotNode[] {
  while (nodes[0]?.field === 'values') nodes = nodes[0].children
  return nodes
}

/** Each path of an axis's tree, as its items joined, and the items shown at each level. */
function pathsOf(nodes: readonly PivotNode[], above: number[] = [], out = { paths: new Set<string>(), levels: [] as Set<number>[] }) {
  for (const n of under(nodes)) {
    const path = [...above, n.item]
    out.paths.add(path.join(','))
    ;(out.levels[above.length] ??= new Set()).add(n.item)
    pathsOf(n.children, path, out)
  }
  return out
}

/** The value of each cell as its value field shows it (`showAs`): the calculation's own where none is set. */
export function shownValues(result: PivotResult): PivotValueAt {
  const values = result.table.values ?? []
  const raw = result.value
  if (values.every((d) => (d.showAs ?? 'normal') === 'normal')) return raw
  const axes = [result.rows, result.columns]
  const real = axes.map((a) => a.fields.filter((f): f is number => f !== 'values'))
  const paths = axes.map((a) => pathsOf(a.items))
  // the items of an axis's field under a path of the fields before it, in the order shown
  const orders = [new Map<string, number[]>(), new Map<string, number[]>()]
  const order = (axis: number, prefix: readonly number[]): number[] => {
    const key = prefix.join(',')
    let out = orders[axis].get(key)
    if (!out) {
      let nodes = axes[axis].items
      for (const p of prefix) nodes = under(nodes).find((n) => n.item === p)?.children ?? []
      orders[axis].set(key, (out = under(nodes).map((n) => n.item)))
    }
    return out
  }
  // the base field's axis (0 rows, 1 columns) and its level among the axis's fields
  const baseOf = (d: PivotDataField) => {
    const f = d.baseField
    const axis = f === undefined ? -1 : real.findIndex((fs) => fs.includes(f))
    return axis < 0 ? undefined : { axis, level: real[axis].indexOf(f!) }
  }

  return (rows, columns, v, fn) => {
    const x = raw(rows, columns, v, fn)
    const d = values[v]
    const show = d?.showAs ?? 'normal'
    if (show === 'normal') return x
    const at = (r: readonly number[], c: readonly number[]) => raw(r, c, v, fn)
    const n = x ?? 0
    switch (show) {
      // a share of a total: an empty total as 0 (#DIV/0!); of a parent: empty over an empty parent
      case 'percentOfTotal':
        return ratio(n, at([], []) ?? 0)
      case 'percentOfRow':
        return ratio(n, at(rows, []) ?? 0)
      case 'percentOfColumn':
        return ratio(n, at([], columns) ?? 0)
      case 'percentOfParentRow':
        return ofParent(n, rows.length ? at(rows.slice(0, -1), columns) : x)
      case 'percentOfParentColumn':
        return ofParent(n, columns.length ? at(rows, columns.slice(0, -1)) : x)
      case 'index': {
        // (value × grand total) / (row total × column total), an empty total as 0
        const [r, c, g] = [at(rows, []), at([], columns), at([], [])].map((y) => y ?? 0)
        const e = [n, r, c, g].find(isError)
        if (e) return e
        return ratio(ratio(pivotMultiply(n as number, g as number), r), c)
      }
    }
    // the rest go along the base field's items, at the cells on its level and under it
    const base = baseOf(d)
    if (!base) return NA
    const { axis, level } = base
    // a base item hidden or filtered out: every cell, the totals too (the set show-items)
    if (typeof d.baseItem === 'number' && (show === 'difference' || show === 'percentDiff' || show === 'percent') && !paths[axis].levels[level]?.has(d.baseItem)) return NA
    const path = axis ? columns : rows
    if (path.length <= level) return null
    const moved = (item: number) => {
      const p = path.slice()
      p[level] = item
      return p
    }
    const along = (item: number) => (axis ? at(rows, moved(item)) : at(moved(item), columns))
    const items = order(axis, path.slice(0, level))
    const own = path[level]
    switch (show) {
      case 'percentOfParent':
        return ofParent(n, axis ? at(rows, path.slice(0, level + 1)) : at(path.slice(0, level + 1), columns))
      case 'difference':
      case 'percentDiff':
      case 'percent': {
        // the previous or next item among those with this line
        const lined = (item: number) => paths[axis].paths.has(moved(item).join(','))
        const near = d.baseItem === 'previous' || d.baseItem === 'next' ? items.filter(lined) : []
        const i = near.indexOf(own)
        const b = d.baseItem === 'previous' ? (near[i - 1] ?? own) : d.baseItem === 'next' ? (near[i + 1] ?? own) : (d.baseItem ?? 0)
        if (b === own) return show !== 'percent' ? null : typeof x === 'number' ? 1 : x
        if (!lined(b)) return NA
        const y = along(b)
        if (show === 'difference') {
          const m = y ?? 0
          return isError(n) ? n : isError(m) ? m : pivotAdd(n as number, -(m as number))
        }
        if (x === null) return NULL
        return show === 'percent' ? ratio(x, y) : ratio(typeof x === 'number' && typeof y === 'number' ? pivotAdd(x, -y) : x, y)
      }
      case 'runTotal':
      case 'percentOfRunningTotal': {
        // the sum of the items up to this one, and of all of them; empty while no item has a number
        let sum: CellValue = 0
        let total: CellValue = null
        let passed = false
        for (const item of items) {
          const y = along(item)
          if (isError(y)) total = y
          else if (typeof y === 'number' && !isError(total)) total = pivotAdd(total ?? 0, y)
          if (!passed) sum = total ?? 0
          if (item === own) {
            if (show === 'runTotal') return sum
            passed = true
          }
        }
        const denominator = path.length === level + 1 ? (axis ? at(rows, path.slice(0, level)) : at(path.slice(0, level), columns)) : total
        return ratio(sum, denominator)
      }
      case 'rankAscending':
      case 'rankDescending': {
        // ties share a rank, and the next value takes the next rank (1, 2, 2, 3: the set show-held)
        if (typeof x !== 'number') return null
        const before = new Set<number>()
        for (const item of items) {
          const y = along(item)
          if (typeof y === 'number' && (show === 'rankAscending' ? y < x : y > x)) before.add(y)
        }
        return before.size + 1
      }
    }
    return x
  }
}
