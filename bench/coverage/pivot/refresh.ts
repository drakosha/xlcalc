/**
 * The set `refresh`: the source edited and the pivot table refreshed: new items, items gone (kept in the cache by
 * `missingItemsLimit`: Excel's default, none, the most), a hidden item gone and back, values and types changed, an
 * error appearing; a source table grown by a row typed below it and resized, a defined name given a smaller range, a
 * range that does not grow; the source edited and, before any refresh, a filter, a sort, the layout, a field moved, a
 * data field added or its function changed (what the pivot table shows from its cache), then refreshed; two pivot
 * tables of one cache, one growing over the other on a refresh, and over cells with values; a cache to refresh on
 * opening; a book Excel saved opened as it is and refreshed; GETPIVOTDATA over a pivot table that a refresh turned into
 * #SPILL!, before and after, and after the cells in its way are cleared and it is refreshed again.
 */
import type { PivotCase, PivotOp } from './cases.js'
import { salesColumns } from './data.js'

/** The rows of `D` where Region is East. */
const EAST = salesColumns()[0].values.flatMap((v, i) => (v?.value === 'East' ? [i + 2] : []))

export function buildPivotCases(): PivotCase[] {
  const base: PivotOp[] = [{ rows: ['Region'] }, { value: { field: 'Price' } }]
  const eastGone: PivotOp[] = EAST.map((r): PivotOp => ({ edit: `D!A${r}`, to: 'West' }))
  const newRegions = (n: number, from = 2): PivotOp[] => Array.from({ length: n }, (_, i): PivotOp => ({ edit: `D!A${from + i}`, to: `Region ${i + 1}` }))
  // GETPIVOTDATA over PT1 (the total, an item, an item it will gain) and PT2, typed before the refresh
  const spillReads: PivotOp[] = [
    { edit: 'G!A1', to: { formula: 'GETPIVOTDATA("Sum of Price",P!$A$3)' } },
    { edit: 'G!A2', to: { formula: 'GETPIVOTDATA("Sum of Price",P!$A$3,"Region","East")' } },
    { edit: 'G!A3', to: { formula: 'GETPIVOTDATA("Sum of Price",P!$A$3,"Region","Region 1")' } },
    { edit: 'G!A4', to: { formula: 'GETPIVOTDATA("Sum of Qty",P!$A$12)' } },
    { edit: 'G!A5', to: { formula: 'GETPIVOTDATA("Sum of Qty",P!$A$12,"Flag",TRUE)' } },
  ]
  const before = (id: string, change: PivotOp[], start: PivotOp[] = base): PivotCase => ({ id, ops: [...start, { read: true }, { edit: 'D!E3', to: 1000 }, { edit: 'D!A5', to: 'Central' }, ...change, { read: true }, { refresh: true }] })
  return [
    { id: 'new-item', ops: [...base, { read: true }, { edit: 'D!A2', to: 'Central' }, { refresh: true }] },
    { id: 'new-items-sorted', ops: [...base, { sort: 'Region', order: 'ascending' }, { read: true }, { edit: 'D!A2', to: 'Central' }, { edit: 'D!A3', to: 'aaa' }, { refresh: true }] },
    { id: 'gone-default', ops: [...base, { read: true }, ...eastGone, { refresh: true }] },
    { id: 'gone-none', ops: [...base, { cache: { missingItems: 'none' } }, { read: true }, ...eastGone, { refresh: true }] },
    { id: 'gone-max', ops: [...base, { cache: { missingItems: 'max' } }, { read: true }, ...eastGone, { refresh: true }] },
    { id: 'gone-hidden-back', ops: [...base, { hide: 'Region', items: ['East'] }, { read: true }, ...eastGone, { refresh: true }, { read: true }, ...EAST.map((r): PivotOp => ({ edit: `D!A${r}`, to: 'East' })), { refresh: true }] },
    { id: 'gone-page', ops: [{ pages: ['Region'] }, { rows: ['Product'] }, { value: { field: 'Price' } }, { page: 'Region', items: ['East'] }, { read: true }, ...eastGone, { refresh: true }] },
    { id: 'values', ops: [...base, { read: true }, { edit: 'D!E2', to: 1000 }, { edit: 'D!E3', to: null }, { refresh: true }] },
    { id: 'types', ops: [{ rows: ['Region'] }, { value: { field: 'Qty' } }, { value: { field: 'Qty', fn: 'countNums' } }, { read: true }, { edit: 'D!F2', to: 'many' }, { edit: 'D!F3', to: '5' }, { refresh: true }] },
    { id: 'error-appears', ops: [...base, { read: true }, { edit: 'D!E2', to: { formula: 'NA()' } }, { refresh: true }] },
    { id: 'item-type', ops: [{ rows: ['Code'] }, { value: { field: 'Qty' } }, { read: true }, { edit: 'D!G2', to: '10' }, { edit: 'D!G3', to: 40 }, { refresh: true }] },
    { id: 'table-grows', source: 'table', ops: [...base, { read: true }, { edit: 'D!A50', to: 'Central' }, { edit: 'D!E50', to: 5 }, { refresh: true }] },
    { id: 'table-resized', source: 'table', ops: [...base, { read: true }, { edit: 'D!A50', to: 'Central' }, { edit: 'D!E50', to: 5 }, { resize: 'Sales', range: 'A1:J50' }, { refresh: true }] },
    { id: 'table-shrinks', source: 'table', ops: [...base, { read: true }, { resize: 'Sales', range: 'A1:J20' }, { refresh: true }] },
    { id: 'name-smaller', source: 'name', ops: [...base, { read: true }, { define: 'SalesData', ref: 'D!$A$1:$J$20' }, { refresh: true }] },
    { id: 'range-fixed', ops: [...base, { read: true }, { edit: 'D!A50', to: 'Central' }, { edit: 'D!E50', to: 5 }, { refresh: true }] },
    { id: 'header-renamed', ops: [...base, { read: true }, { edit: 'D!E1', to: 'Cost' }, { refresh: true }] },
    { id: 'header-renamed-row', ops: [...base, { read: true }, { edit: 'D!A1', to: 'Area' }, { refresh: true }] },
    before('stale-hide', [{ hide: 'Region', items: ['South'] }]),
    before('stale-filter', [{ filter: 'Region', type: 'ValueIsGreaterThan', value1: 250, by: 'Sum of Price' }]),
    before('stale-sort', [{ sort: 'Region', order: 'descending', by: 'Sum of Price' }]),
    before('stale-layout', [{ layout: 'tabular' }], [{ rows: ['Region', 'Flag'] }, { value: { field: 'Price' } }]),
    before('stale-move', [{ move: 'Region', to: 'column' }]),
    before('stale-new-field', [{ rows: ['Flag'] }]),
    before('stale-new-value', [{ value: { field: 'Price', fn: 'max' } }]),
    before('stale-function', [{ value: { field: 'Qty' } }]),
    {
      id: 'two-overlap',
      ops: [...base, { create: 'PT2', at: 'P!A12' }, { rows: ['Flag'], pt: 'PT2' }, { value: { field: 'Qty' }, pt: 'PT2' }, { read: true }, ...newRegions(5), { refresh: true }],
    },
    {
      id: 'two-overlap-all',
      ops: [...base, { create: 'PT2', at: 'P!A12' }, { rows: ['Flag'], pt: 'PT2' }, { value: { field: 'Qty' }, pt: 'PT2' }, { read: true }, ...newRegions(5), { refreshAll: true }],
    },
    {
      id: 'two-side',
      ops: [...base, { create: 'PT2', at: 'P!D3' }, { rows: ['Flag'], pt: 'PT2' }, { value: { field: 'Qty' }, pt: 'PT2' }, { read: true }, ...newRegions(3), { refresh: true, pt: 'PT2' }],
    },
    {
      id: 'two-caches',
      ops: [...base, { create: 'PT2', at: 'P!D3', source: 'range' }, { rows: ['Region'], pt: 'PT2' }, { value: { field: 'Price' }, pt: 'PT2' }, { read: true }, { edit: 'D!A2', to: 'Central' }, { refresh: true }],
    },
    { id: 'over-values', ops: [...base, { edit: 'P!A12', to: 'kept?' }, { edit: 'P!B13', to: 7 }, { read: true }, ...newRegions(5), { refresh: true }], watch: ['P!A12:B16'] },
    {
      id: 'two-overlap-gpd',
      ops: [...base, { create: 'PT2', at: 'P!A12' }, { rows: ['Flag'], pt: 'PT2' }, { value: { field: 'Qty' }, pt: 'PT2' }, ...spillReads, { read: true }, ...newRegions(5), { refresh: true }],
      watch: ['G!A1:A5'],
    },
    { id: 'over-values-gpd', ops: [...base, { edit: 'P!A12', to: 'kept?' }, ...spillReads.slice(0, 3), { read: true }, ...newRegions(5), { refresh: true }, { read: true }, { edit: 'P!A12', to: null }, { refresh: true }], watch: ['G!A1:A3'] },
    { id: 'refresh-on-open', ops: [...base, { cache: { refreshOnLoad: true } }, { edit: 'D!A2', to: 'Central' }] },
    { id: 'open-saved', file: 'test/excel/win/pivot/basic/columns.xlsx', ops: [{ read: true }, { refreshAll: true }] },
    { id: 'open-saved-two', file: 'test/excel/win/pivot/basic/page-one.xlsx', ops: [{ read: true }, { refreshAll: true }] },
  ]
}
