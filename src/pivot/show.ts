/**
 * Show Values As: a value field's cell as a share of a total or of its parent, set against a base item (a share, a
 * difference), a running total over the base field's items, a rank among them, or an index, calculated over the
 * values of the cube (`PivotResult.value`, which answers for any items). The layout takes its values from here.
 */
import type { CellValue, ErrorCode, PivotDataField, PivotFunction } from '../contract.js'
import type { PivotNode, PivotResult } from './result.js'

/** A value at the paths of the rows and the columns, of a value field, by a subtotal's own function. */
export type PivotValueAt = (rows: readonly number[], columns: readonly number[], value: number, fn?: PivotFunction) => CellValue

const NA: CellValue = { error: 'NA' }
const DIV0: CellValue = { error: 'DIV0' }

const isError = (x: CellValue): x is { error: ErrorCode } => x !== null && typeof x === 'object'

/** `x / y`: the error of `y`, #DIV/0! by zero, #N/A without a value to divide by. */
function ratio(x: number, y: CellValue): CellValue {
  if (isError(y)) return y
  if (typeof y !== 'number') return NA
  return y === 0 ? DIV0 : x / y
}

/** The items under a node's path, `values` passed over (its copies under each value field are alike). */
function under(nodes: readonly PivotNode[]): readonly PivotNode[] {
  while (nodes[0]?.field === 'values') nodes = nodes[0].children
  return nodes
}

/** The value of each cell as its value field shows it (`showAs`): the calculation's own where none is set. */
export function shownValues(result: PivotResult): PivotValueAt {
  const values = result.table.values ?? []
  const raw = result.value
  if (values.every((d) => (d.showAs ?? 'normal') === 'normal')) return raw
  const axes = [result.rows, result.columns]
  const real = axes.map((a) => a.fields.filter((f): f is number => f !== 'values'))
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
    const f = d.baseField ?? 0
    const axis = real.findIndex((fs) => fs.includes(f))
    return axis < 0 ? undefined : { axis, level: real[axis].indexOf(f) }
  }

  return (rows, columns, v, fn) => {
    const x = raw(rows, columns, v, fn)
    const d = values[v]
    const show = d?.showAs ?? 'normal'
    if (show === 'normal' || typeof x !== 'number') return x
    const at = (r: readonly number[], c: readonly number[]) => raw(r, c, v, fn)
    switch (show) {
      case 'percentOfTotal':
        return ratio(x, at([], []))
      case 'percentOfRow':
        return ratio(x, at(rows, []))
      case 'percentOfColumn':
        return ratio(x, at([], columns))
      case 'percentOfParentRow':
        return rows.length ? ratio(x, at(rows.slice(0, -1), columns)) : ratio(x, x)
      case 'percentOfParentColumn':
        return columns.length ? ratio(x, at(rows, columns.slice(0, -1))) : ratio(x, x)
      case 'index': {
        // (value × grand total) / (row total × column total)
        const [r, c, g] = [at(rows, []), at([], columns), at([], [])]
        for (const y of [r, c, g]) if (isError(y)) return y
        if (typeof r !== 'number' || typeof c !== 'number' || typeof g !== 'number') return NA
        return ratio(x * g, r * c)
      }
    }
    // the rest go along the base field's items, at the cells on its level and under it
    const base = baseOf(d)
    if (!base) return NA
    const { axis, level } = base
    const path = axis ? columns : rows
    if (path.length <= level) return null
    const along = (item: number) => {
      const p = path.slice()
      p[level] = item
      return axis ? at(rows, p) : at(p, columns)
    }
    const items = order(axis, path.slice(0, level))
    const own = path[level]
    switch (show) {
      case 'percentOfParent':
        return ratio(x, axis ? at(rows, path.slice(0, level + 1)) : at(path.slice(0, level + 1), columns))
      case 'difference':
      case 'percentDiff':
      case 'percent': {
        const i = items.indexOf(own)
        const b = d.baseItem === 'previous' ? (items[i - 1] ?? own) : d.baseItem === 'next' ? (items[i + 1] ?? own) : (d.baseItem ?? 0)
        if (b === own) return show === 'percent' ? 1 : null
        const y = along(b)
        if (typeof y !== 'number') return y === null ? NA : y
        return show === 'difference' ? x - y : show === 'percent' ? ratio(x, y) : ratio(x - y, y)
      }
      case 'runTotal':
      case 'percentOfRunningTotal': {
        let sum = 0
        let total = 0
        let passed = false
        for (const item of items) {
          const y = along(item)
          if (isError(y)) return y
          if (typeof y === 'number') {
            total += y
            if (!passed) sum += y
          }
          if (item === own) {
            if (show === 'runTotal') return sum
            passed = true
          }
        }
        return ratio(sum, total)
      }
      case 'rankAscending':
      case 'rankDescending': {
        const up = show === 'rankAscending'
        let rank = 1
        for (const item of items) {
          const y = along(item)
          if (typeof y === 'number' && (up ? y < x : y > x)) rank++
        }
        return rank
      }
    }
    return x
  }
}
