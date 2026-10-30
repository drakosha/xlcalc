/**
 * A PivotTable's calculation as Excel saved it, for the tests of the layout: the trees of the axes from the saved lines
 * (`rowLines`, `columnLines`) and the values from the cells of the area, so that laying the PivotTable out again should
 * give back the lines, the place and the cells Excel saved. Not the calculation: that is `computePivot`'s.
 */
import type { CellInput, PivotCache, PivotFunction, PivotLine, PivotTable, SheetInput } from '../src/contract.js'
import { ERROR_CODE } from '../src/syntax/errors.js'
import { formatAddress, parseAddress } from '../src/model.js'
import type { PivotAxisResult, PivotNode, PivotResult } from '../src/pivot/result.js'

/** A saved line in full: the items of each field from the first (`values` holding its value field), its type and value field. */
interface Full {
  items: number[]
  type?: PivotLine['type']
  value: number
}

function fullLines(lines: PivotLine[] | undefined): Full[] {
  let last: number[] = []
  return (lines ?? []).map((l) => {
    const items = l.type === 'grand' ? [] : [...last.slice(0, l.repeat ?? 0), ...l.items]
    if (l.type !== 'grand') last = items
    return { items, ...(l.type ? { type: l.type } : {}), value: l.value ?? 0 }
  })
}

/** The tree of an axis from its data lines, items in the order of their first line. */
function treeOf(fields: PivotAxisResult['fields'], lines: Full[]): PivotNode[] {
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

const key = (path: readonly number[], v: number, fn: PivotFunction | undefined) => `${path.join(',')}|${v}|${fn ?? ''}`

/** The value of a cell as the calculation gives it: errors by their code, an empty cell null. */
function valueOf(c: CellInput | undefined) {
  if (!c) return null
  if (c.formula !== undefined && ERROR_CODE[c.formula]) return { error: ERROR_CODE[c.formula] }
  const v = c.value
  return v === undefined ? null : typeof v === 'object' && v !== null ? null : v
}

/** The calculation that the saved layout of `table` on `sheet` shows. */
export function savedResult(table: PivotTable, cache: PivotCache, sheet: SheetInput): PivotResult {
  const rowFields = table.rows ?? []
  const colFields = table.columns ?? []
  const rows = fullLines(table.rowLines)
  const cols = fullLines(table.columnLines)
  const at = parseAddress(table.location.range.split(':')[0])
  const top = at.row + table.location.firstDataRow
  const left = at.col + table.location.firstDataColumn
  // each line's path (the items of the fields but `values`), value field and subtotal function, to its index
  const index = (fields: PivotAxisResult['fields'], lines: Full[]) => {
    const m = new Map<string, number>()
    const vAt = fields.indexOf('values')
    lines.forEach((l, i) => {
      // a blank line, and a value field's line over the items under it, show no values
      if (l.type === 'blank' || (!l.type && l.items.length - 1 === vAt && vAt < fields.length - 1)) return
      const path = l.items.filter((_, k) => k !== vAt)
      const fn = l.type && l.type !== 'subtotal' && l.type !== 'grand' ? l.type : undefined
      const v = vAt >= 0 && vAt < l.items.length ? l.items[vAt] : l.value
      for (const vv of vAt >= 0 && vAt >= l.items.length && l.type === undefined ? [] : [v]) {
        const k = key(path, vv, fn)
        if (!m.has(k)) m.set(k, i)
      }
    })
    return m
  }
  const rowAt = index(rowFields, rows)
  const colAt = index(colFields, cols)
  const pages = (table.pages ?? []).map((p) => ({ field: p.field, items: p.item === undefined ? [] : [p.item], all: p.item === undefined }))
  return {
    table,
    cache,
    rows: { fields: rowFields, items: treeOf(rowFields, rows) },
    columns: { fields: colFields, items: treeOf(colFields, cols) },
    pages,
    value(r, c, v, fn) {
      const i = rowAt.get(key(r, rowFields.includes('values') ? v : 0, fn)) ?? rowAt.get(key(r, 0, fn))
      const j = colAt.get(key(c, colFields.includes('values') ? v : 0, fn)) ?? colAt.get(key(c, 0, fn))
      if (i === undefined || j === undefined) return null
      return valueOf(sheet.cells[formatAddress(top + i, left + j)])
    },
    hasData: () => true,
  }
}
