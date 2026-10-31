/**
 * The set `alerts-clear`: what Excel clears of a pivot table and what it deletes, over the pivot table of the set
 * `alerts` (`P!A1:B1` its Filters area, `A3:D10` its table), with `x` in the blank row under the filters (`A2`), a
 * number right of it (`F5`) and below it (`B12`), and a second pivot table at `H3` over the same cache in some cases:
 * contents cleared (`ClearContents`) over `TableRange2`, over the table alone, over the table and the Filters area
 * without the blank row between (two areas), over the pivot table with cells around it, its columns, the whole sheet,
 * both pivot tables, one and part of the other; everything cleared (`Clear`) and the formats alone (`ClearFormats`).
 * The cells `A1:L14` are read after each.
 */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const grid: PivotOp[] = [{ pages: ['Product'] }, { rows: ['Region'] }, { columns: ['Flag'] }, { value: { field: 'Qty' } }]
  const beside: PivotOp[] = [{ edit: 'P!A2', to: 'x' }, { edit: 'P!F5', to: 5 }, { edit: 'P!B12', to: 7 }]
  const second: PivotOp[] = [{ create: 'PT2', at: 'P!H3' }, { rows: ['Code'], pt: 'PT2' }, { value: { field: 'Qty' }, pt: 'PT2' }]
  const watch = ['P!A1:L14']
  const at = (id: string, op: PivotOp, two = false): PivotCase => ({ id, watch, ops: [...grid, ...(two ? second : []), ...beside, { read: true }, op] })
  return [
    at('table-range-2', { edit: 'P!A1:D10', to: null }),
    at('table-alone', { edit: 'P!A3:D10', to: null }),
    at('table-and-filters', { edit: 'P!A1:B1,A3:D10', to: null }),
    at('filters-alone', { edit: 'P!A1:B1', to: null }),
    at('with-cells-around', { edit: 'P!A1:F12', to: null }),
    at('its-columns', { edit: 'P!A:D', to: null }),
    at('whole-sheet', { edit: 'P!1:1048576', to: null }),
    at('both', { edit: 'P!A1:L14', to: null }, true),
    at('one-and-part', { edit: 'P!A1:D10,H3:H4', to: null }, true),
    at('one-of-two', { edit: 'P!A1:D10', to: null }, true),
    at('clear-all', { clear: 'P!A1:D10', what: 'all' }),
    at('clear-formats', { clear: 'P!A1:D10', what: 'formats' }),
  ]
}
