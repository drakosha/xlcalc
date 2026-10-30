/**
 * A PivotTable's calculation as Excel saved it, for the tests of the layout: the trees of the axes from the saved lines
 * (`rowLines`, `columnLines`) and the values from the cells of the area, so that laying the PivotTable out again should
 * give back the lines, the place and the cells Excel saved. Not the calculation: that is `computePivot`'s.
 */
import type { CellInput, PivotCache, PivotFunction, PivotTable, SheetInput } from '../src/contract.js'
import { ERROR_CODE } from '../src/syntax/errors.js'
import { formatAddress, parseAddress } from '../src/model.js'
import type { PivotAxisResult, PivotResult } from '../src/pivot/result.js'
import { fullLines, savedResult as savedLayout, type Full } from '../src/pivot/saved.js'

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
        // a subtotal's own line over the line of its item, which shows none when the subtotal goes below
        const k = key(path, vv, fn)
        if (!m.has(k) || l.type) m.set(k, i)
      }
    })
    return m
  }
  const rowAt = index(rowFields, rows)
  const colAt = index(colFields, cols)
  return {
    ...savedLayout(table, cache),
    value(r, c, v, fn) {
      // a line by its path, value field and subtotal function: the function is the line's of one axis only
      const find = (at: Map<string, number>, path: readonly number[], own: boolean) => {
        const vv = own ? v : 0
        return at.get(key(path, vv, fn)) ?? at.get(key(path, vv, undefined))
      }
      const i = find(rowAt, r, rowFields.includes('values'))
      const j = find(colAt, c, colFields.includes('values'))
      if (i === undefined || j === undefined) return null
      return valueOf(sheet.cells[formatAddress(top + i, left + j)])
    },
    hasData: () => true,
  }
}
