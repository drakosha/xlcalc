/**
 * A PivotTable's lines on an axis as Excel lays them out (xlsx `rowItems`, `colItems`): the items in the order shown,
 * each item's subtotals at the top of its group or below it, blank lines after a group, the grand totals; with what the
 * cells of each line show (`AxisLine`).
 */
import type { PivotAxisField, PivotField, PivotFunction, PivotLayout, PivotLine, PivotTable } from '../contract.js'
import type { PivotAxisResult, PivotNode } from './result.js'

/** What a line of an axis shows, besides its saved form. */
export interface AxisLine {
  line: PivotLine
  /** `data`: an item (with its subtotal when it is not the last field's); `item`: an item over its group, values empty; `subtotal`, `grand`, `blank`. */
  kind: 'data' | 'item' | 'subtotal' | 'grand' | 'blank'
  /** The items on the axis's fields up to the line's field, `values` included (its value field's index there). */
  items: number[]
  /** The items of the fields other than `values`: the path of the cube the line's values are at. */
  path: number[]
  /** The value field of the line, when the value fields lie across this axis and the line has one. */
  value?: number
  /** A subtotal's own function. */
  fn?: PivotFunction
  /** The fields on the line from which on it shows a label (`items` from it are new to the line). */
  from: number
}

/** A field's form on the axis: its own, else the PivotTable's. */
export function formOf(table: PivotTable, field: PivotAxisField): PivotLayout {
  return (field === 'values' ? undefined : table.fields[field]?.layout) ?? table.layout ?? 'compact'
}

/** The functions of a field's subtotals: `subtotal` for the automatic one, none for `[]`. */
const functionsOf = (f: PivotField | undefined): ('subtotal' | PivotFunction)[] => (f?.subtotals === undefined ? ['subtotal'] : f.subtotals)

/** Whether an axis has a field other than `values`. */
export const hasItems = (fields: readonly PivotAxisField[]) => fields.some((f) => f !== 'values')

/**
 * The lines of an axis. On rows, an item over a group of a compact or outline field has a line of its own, which shows
 * its subtotal when the subtotal goes at the top; on columns, and in tabular form, a subtotal is a line after the group.
 * A field has subtotals only when a field other than `values` lies under it; with the value fields under it, a subtotal
 * per value field. `grand`: the axis's grand total shown.
 */
export function axisLines(table: PivotTable, axis: PivotAxisResult, rows: boolean, grand: boolean): AxisLine[] {
  const { fields } = axis
  const valuesAt = fields.indexOf('values')
  const out: AxisLine[] = []
  let last: number[] = []
  const push = (kind: AxisLine['kind'], items: number[], extra: { value?: number; fn?: PivotFunction } = {}) => {
    const depth = items.length - 1
    let repeat = 0
    if (kind !== 'grand') while (repeat < depth && repeat < last.length && last[repeat] === items[repeat]) repeat++
    const line: PivotLine = { items: kind === 'grand' ? [0] : items.slice(repeat) }
    if (kind === 'subtotal' || kind === 'grand' || kind === 'blank') line.type = kind === 'subtotal' ? (extra.fn ?? 'subtotal') : kind
    if (repeat) line.repeat = repeat
    const value = extra.value ?? (valuesAt >= 0 && valuesAt <= depth ? items[valuesAt] : undefined)
    if (value) line.value = value
    const path = items.filter((_, k) => k !== valuesAt)
    const a: AxisLine = { line, kind, items, path, from: repeat }
    if (value !== undefined) a.value = value
    if (extra.fn) a.fn = extra.fn
    out.push(a)
    last = items
  }
  // the value fields a subtotal at depth `d` is made for: each, when they lie under it; the line's own otherwise
  const subtotalValues = (d: number, items: number[]): (number | undefined)[] =>
    valuesAt > d ? (table.values ?? []).map((_, v) => v) : [valuesAt >= 0 ? items[valuesAt] : undefined]
  const walk = (nodes: PivotNode[], above: number[]) => {
    for (const n of nodes) {
      const d = above.length
      const items = [...above, n.item]
      const leaf = d === fields.length - 1
      const f = n.field === 'values' ? undefined : table.fields[n.field]
      const sums = n.field !== 'values' && fields.slice(d + 1).some((x) => x !== 'values') ? functionsOf(f) : []
      const each = sums.length ? subtotalValues(d, items) : []
      const form = formOf(table, n.field)
      const collapsed = !leaf && n.field !== 'values' && table.fields[n.field]?.items?.[n.item]?.collapsed === true
      const top = rows && form !== 'tabular' && f?.subtotalTop !== false && sums.length === 1 && sums[0] === 'subtotal' && each.length === 1
      if (leaf || collapsed) push('data', items)
      else {
        if (rows && form !== 'tabular') push(top ? 'data' : 'item', items)
        walk(n.children, items)
      }
      if (!leaf && !(top || (collapsed && each.length === 1 && sums.length === 1)))
        for (const fn of sums) for (const value of each) push('subtotal', items, { ...(value === undefined ? {} : { value }), ...(fn === 'subtotal' ? {} : { fn }) })
      if (rows && !leaf && f?.blankLine) push('blank', items)
    }
  }
  walk(axis.items, [])
  if (grand && hasItems(fields)) {
    const each = valuesAt >= 0 ? (table.values ?? []).map((_, v) => v) : [undefined]
    for (const value of each) push('grand', [], value === undefined ? {} : { value })
  }
  if (!out.length) push('data', [])
  return out
}
