/**
 * The held-out set `slicer-held`, over the rows of `stock`: a slicer over a page field and over a field on no axis, over
 * two pivot tables of a cache, its items sorted and its items without data placed or hidden, a pivot table connected
 * and disconnected; a timeline by levels and over one day.
 */
import { serial } from './data.js'
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const base: PivotOp[] = [{ rows: ['Region'] }, { value: { field: 'Qty' } }]
  const two: PivotOp[] = [...base, { create: 'PT2', at: 'P!E3' }, { rows: ['Product'], pt: 'PT2' }, { value: { field: 'Price' }, pt: 'PT2' }]
  const c = (id: string, ops: PivotOp[]): PivotCase => ({ id, data: 'stock', ops })
  return [
    c('page-field', [{ pages: ['Product'] }, ...base, { slicer: 'Product', name: 'Slicer_Product' }, { select: 'Slicer_Product', items: ['Fig', 'Pear'] }]),
    c('page-field-one', [{ pages: ['Product'] }, ...base, { page: 'Product', items: ['Plum'] }, { slicer: 'Product', name: 'Slicer_Product' }]),
    c('off-axes', [...base, { slicer: 'Code', name: 'Slicer_Code' }, { select: 'Slicer_Code', items: ['25'] }]),
    c('off-axes-text-number', [...base, { slicer: 'Code', name: 'Slicer_Code' }, { select: 'Slicer_Code', items: ['15'] }]),
    c('two-pivots', [...two, { slicer: 'Flag', name: 'Slicer_Flag', pivots: ['PT2'] }, { select: 'Slicer_Flag', items: ['FALSE'] }]),
    c('two-slicers-cross', [...two, { slicer: 'Region', name: 'Slicer_Region', pivots: ['PT2'] }, { slicer: 'Product', name: 'Slicer_Product', pivots: ['PT2'] }, { select: 'Slicer_Product', items: ['7'] }]),
    c('sort-descending', [...base, { slicer: 'Region', name: 'Slicer_Region' }, { slicerSettings: 'Slicer_Region', order: 'descending' }, { select: 'Slicer_Region', items: ['North', 'east'] }]),
    c('sort-source', [...base, { slicer: 'Product', name: 'Slicer_Product' }, { slicerSettings: 'Slicer_Product', order: 'source' }]),
    c('no-data-hidden', [...base, { slicer: 'Region', name: 'Slicer_Region' }, { slicer: 'Product', name: 'Slicer_Product' }, { slicerSettings: 'Slicer_Region', noData: 'hidden' }, { select: 'Slicer_Product', items: ['7'] }]),
    c('no-data-none', [...base, { slicer: 'Region', name: 'Slicer_Region' }, { slicer: 'Product', name: 'Slicer_Product' }, { slicerSettings: 'Slicer_Region', noData: 'none' }, { select: 'Slicer_Product', items: ['7'] }]),
    c('no-data-dimmed', [...base, { slicer: 'Region', name: 'Slicer_Region' }, { slicer: 'Product', name: 'Slicer_Product' }, { slicerSettings: 'Slicer_Region', noData: 'dimmed' }, { select: 'Slicer_Product', items: ['7'] }]),
    c('connect', [...two, { slicer: 'Region', name: 'Slicer_Region' }, { select: 'Slicer_Region', items: ['South'] }, { read: true }, { link: 'Slicer_Region', pivot: 'PT2' }]),
    c('disconnect', [...two, { slicer: 'Region', name: 'Slicer_Region', pivots: ['PT2'] }, { select: 'Slicer_Region', items: ['South'] }, { read: true }, { link: 'Slicer_Region', pivot: 'PT2', off: true }, { read: true }, { select: 'Slicer_Region', items: ['West'] }]),
    c('timeline-years', [...base, { timeline: 'Date', name: 'Timeline_Date', from: serial(2025, 1, 1), to: serial(2025, 12, 31), level: 'years' }]),
    c('timeline-months-two', [...two, { timeline: 'Date', name: 'Timeline_Date', from: serial(2024, 2, 1), to: serial(2024, 8, 31), level: 'months' }, { link: 'Timeline_Date', pivot: 'PT2' }]),
    c('timeline-one-day', [{ rows: ['Date'] }, { value: { field: 'Qty' } }, { timeline: 'Date', name: 'Timeline_Date', from: serial(2025, 4, 22), to: serial(2025, 4, 22), level: 'days' }]),
  ]
}
