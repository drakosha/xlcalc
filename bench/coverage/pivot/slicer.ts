/**
 * The set `slicer`: a slicer on a row field and on a field the pivot table does not show, one item selected and several,
 * all; a slicer over two pivot tables of one cache; items with no data under another slicer; a timeline on dates with a
 * range of dates, its level, over two pivot tables. Each phase reads the slicers' items (selected, with data).
 */
import { serial } from './data.js'
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const base: PivotOp[] = [{ rows: ['Region'] }, { value: { field: 'Price' } }]
  const two: PivotOp[] = [...base, { create: 'PT2', at: 'P!D3' }, { rows: ['Product'], pt: 'PT2' }, { value: { field: 'Qty' }, pt: 'PT2' }]
  return [
    { id: 'row-one', ops: [...base, { slicer: 'Region', name: 'Slicer_Region' }, { select: 'Slicer_Region', items: ['West'] }] },
    { id: 'row-two', ops: [...base, { slicer: 'Region', name: 'Slicer_Region' }, { select: 'Slicer_Region', items: ['West', 'North'] }] },
    { id: 'row-blank', ops: [...base, { slicer: 'Region', name: 'Slicer_Region' }, { select: 'Slicer_Region', items: ['(blank)'] }] },
    { id: 'row-all', ops: [...base, { slicer: 'Region', name: 'Slicer_Region' }, { select: 'Slicer_Region', items: ['West'] }, { read: true }, { select: 'Slicer_Region', items: ['East', 'North', 'South', 'West', '(blank)'] }] },
    { id: 'other-field', ops: [...base, { slicer: 'Flag', name: 'Slicer_Flag' }, { select: 'Slicer_Flag', items: ['TRUE'] }] },
    { id: 'other-field-numbers', ops: [...base, { slicer: 'Code', name: 'Slicer_Code' }, { select: 'Slicer_Code', items: ['10', '30'] }] },
    { id: 'two-pivots', ops: [...two, { slicer: 'Region', name: 'Slicer_Region', pivots: ['PT2'] }, { select: 'Slicer_Region', items: ['South'] }] },
    { id: 'two-pivots-one', ops: [...two, { slicer: 'Region', name: 'Slicer_Region' }, { select: 'Slicer_Region', items: ['South'] }] },
    { id: 'no-data', ops: [...base, { slicer: 'Region', name: 'Slicer_Region' }, { slicer: 'Product', name: 'Slicer_Product' }, { select: 'Slicer_Product', items: ['100'] }] },
    { id: 'with-hidden', ops: [...base, { hide: 'Region', items: ['East'] }, { slicer: 'Region', name: 'Slicer_Region' }] },
    { id: 'page-and-slicer', ops: [{ pages: ['Flag'] }, ...base, { page: 'Flag', items: ['FALSE'] }, { slicer: 'Region', name: 'Slicer_Region' }, { select: 'Slicer_Region', items: ['East', 'West'] }] },
    { id: 'timeline-range', ops: [...base, { timeline: 'Date', name: 'Timeline_Date', from: serial(2024, 3, 1), to: serial(2024, 12, 31) }] },
    { id: 'timeline-level', ops: [...base, { timeline: 'Date', name: 'Timeline_Date', from: serial(2025, 1, 1), to: serial(2025, 6, 30), level: 'quarters' }] },
    { id: 'timeline-days', ops: [...base, { timeline: 'Date', name: 'Timeline_Date', from: serial(2024, 1, 1), to: serial(2024, 1, 1), level: 'days' }] },
    { id: 'timeline-rows-dates', ops: [{ rows: ['Date'] }, { value: { field: 'Qty' } }, { timeline: 'Date', name: 'Timeline_Date', from: serial(2024, 5, 1), to: serial(2025, 2, 28), level: 'months' }] },
    { id: 'timeline-and-slicer', ops: [...base, { timeline: 'Date', name: 'Timeline_Date', from: serial(2025, 1, 1), to: serial(2025, 12, 31) }, { slicer: 'Product', name: 'Slicer_Product' }, { select: 'Slicer_Product', items: ['Apple'] }] },
  ]
}
