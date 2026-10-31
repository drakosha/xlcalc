/**
 * A PivotTable as its saved layout shows it (`rowLines`, `columnLines`): the items of its axes as trees and the items of
 * its Filters area, without values, for what needs the parts of its cells but not the calculation (an entry into a
 * cell, `pivotEntry`).
 */
import type { PivotCache, PivotLine, PivotTable } from '../contract.js'
import type { PivotAxisResult, PivotNode, PivotPageResult, PivotResult } from './result.js'

/** A saved line in full: the items of each field from the first (`values` holding its value field), its type and value field. */
export interface Full {
  items: number[]
  type?: PivotLine['type']
  value: number
}

export function fullLines(lines: PivotLine[] | undefined): Full[] {
  let last: number[] = []
  return (lines ?? []).map((l) => {
    const items = l.type === 'grand' ? [] : [...last.slice(0, l.repeat ?? 0), ...l.items]
    if (l.type !== 'grand') last = items
    return { items, ...(l.type ? { type: l.type } : {}), value: l.value ?? 0 }
  })
}

/** The tree of an axis from its data lines, items in the order of their first line. */
export function treeOf(fields: PivotAxisResult['fields'], lines: Full[]): PivotNode[] {
  const top: PivotNode[] = []
  for (const l of lines) {
    if (l.type) continue
    let level = top
    l.items.forEach((item, k) => {
      let n = level.find((x) => x.item === item)
      if (!n) level.push((n = { field: fields[k], item, children: [] }))
      level = n.children
    })
  }
  return top
}

/** The Filters area's fields: the item shown, else the items the field does not hide. */
export function savedPages(table: PivotTable): PivotPageResult[] {
  return (table.pages ?? []).map((p) => {
    const items = p.item !== undefined ? [p.item] : (table.fields[p.field].items ?? []).flatMap((it, i) => (it.hidden ? [] : [i]))
    return { field: p.field, items, all: p.item === undefined && !table.fields[p.field].items?.some((it) => it.hidden) }
  })
}

/** The PivotTable's saved layout as a calculation without values. */
export function savedResult(table: PivotTable, cache: PivotCache): PivotResult {
  const axis = (fields: PivotAxisResult['fields'] = [], lines?: PivotLine[]) => ({ fields, items: treeOf(fields, fullLines(lines)) })
  return { table, cache, rows: axis(table.rows, table.rowLines), columns: axis(table.columns, table.columnLines), pages: savedPages(table), value: () => null, hasData: () => false }
}
