/**
 * The held-out set `group-held`, over the rows of `stock`: dates grouped by levels, alone and together, with starts,
 * ends and numbers of days; times by hours and minutes; whole numbers by intervals with their edges; prices made all
 * numbers and grouped; items grouped by hand over two fields at once and on the columns; grouping on one of two pivot
 * tables of a cache, seen in the other; a group refreshed with new dates.
 */
import { serial } from './data.js'
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const dates: PivotOp[] = [{ rows: ['Date'] }, { value: { field: 'Price' } }]
  const c = (id: string, ops: PivotOp[]): PivotCase => ({ id, data: 'stock', ops })
  return [
    c('years-quarters', [...dates, { group: 'Date', periods: ['quarters', 'years'] }]),
    c('months-days', [...dates, { group: 'Date', periods: ['days', 'months'] }]),
    c('quarters-months', [...dates, { group: 'Date', periods: ['months', 'quarters'] }]),
    c('days-14-end', [...dates, { group: 'Date', periods: ['days'], by: 14, end: { d: serial(2025, 6, 30) } }]),
    c('years-start-end', [...dates, { group: 'Date', periods: ['years'], start: { d: serial(2024, 6, 1) }, end: { d: serial(2025, 3, 1) } }]),
    c('dates-columns', [{ rows: ['Region'] }, { columns: ['Date'] }, { value: { field: 'Qty' } }, { group: 'Date', periods: ['years'] }]),
    c('when-hours-minutes', [{ rows: ['When'] }, { value: { field: 'Qty' } }, { group: 'When', periods: ['minutes', 'hours'] }]),
    c('when-days-hours', [{ rows: ['When'] }, { value: { field: 'Qty' } }, { group: 'When', periods: ['hours', 'days'] }]),
    c('qty-edges', [{ rows: ['Qty'] }, { value: { field: 'Price' } }, { group: 'Qty', start: 3, end: 7, by: 2 }]),
    c('qty-zero-4', [{ rows: ['Qty'] }, { value: { field: 'Price' } }, { group: 'Qty', start: 0, by: 4 }]),
    c('qty-end-beyond', [{ rows: ['Qty'] }, { value: { field: 'Price' } }, { group: 'Qty', start: 1, end: 20, by: 5 }]),
    c('price-numbers', [{ edit: 'D!E14', to: 12.5 }, { refresh: true }, { rows: ['Price'] }, { value: { field: 'Qty' } }, { group: 'Price', start: 10, end: 40, by: 5 }]),
    c('items-two-fields', [{ rows: ['Region', 'Product'] }, { value: { field: 'Qty' } }, { groupItems: 'Region', items: ['Central', 'North'] }, { groupItems: 'Product', items: ['Plum', 'Pear'] }]),
    c('items-three', [{ rows: ['Region'] }, { value: { field: 'Qty' } }, { groupItems: 'Region', items: ['West', 'east', 'South'] }]),
    c('items-columns-rows', [{ rows: ['Product'] }, { columns: ['Region'] }, { value: { field: 'Qty' } }, { groupItems: 'Region', items: ['South', 'North'] }, { groupItems: 'Product', items: ['Apple', 'Fig'] }]),
    c('shared-dates', [...dates, { create: 'PT2', at: 'P!E3' }, { rows: ['Date'], pt: 'PT2' }, { value: { field: 'Qty' }, pt: 'PT2' }, { group: 'Date', periods: ['quarters'], pt: 'PT2' }]),
    c('shared-items', [{ rows: ['Region'] }, { value: { field: 'Qty' } }, { create: 'PT2', at: 'P!E3' }, { rows: ['Region'], pt: 'PT2' }, { columns: ['Flag'], pt: 'PT2' }, { value: { field: 'Price' }, pt: 'PT2' }, { groupItems: 'Region', items: ['Central', 'West'], pt: 'PT2' }]),
    c('refresh-new-dates', [...dates, { group: 'Date', periods: ['quarters', 'years'] }, { read: true }, { edit: 'D!C2', to: serial(2023, 11, 5) }, { edit: 'D!C3', to: serial(2026, 4, 9) }, { refresh: true }]),
    c('ungroup-items', [{ rows: ['Region'] }, { value: { field: 'Qty' } }, { groupItems: 'Region', items: ['Central', 'West'] }, { read: true }, { ungroup: 'Region2' }]),
  ]
}
