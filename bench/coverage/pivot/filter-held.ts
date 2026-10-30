/**
 * The held-out set `filter-held`, over the rows of `stock`: hidden items, page fields, label, value, top 10 and date
 * filters of `filter` on other items and fields, and in other combinations.
 */
import { serial } from './data.js'
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const base: PivotOp[] = [{ rows: ['Region'] }, { value: { field: 'Qty' } }]
  const two: PivotOp[] = [{ rows: ['Product', 'Region'] }, { value: { field: 'Qty' } }]
  const c = (id: string, ops: PivotOp[]): PivotCase => ({ id, data: 'stock', ops })
  return [
    c('hide-two', [...base, { hide: 'Region', items: ['Central', '(blank)'] }]),
    c('hide-lower', [...base, { hide: 'Region', items: ['East'] }]),
    c('hide-outer-and-inner', [...two, { hide: 'Product', items: ['Fig'] }, { hide: 'Region', items: ['South'] }]),
    c('page-several-two-fields', [{ pages: ['Product', 'Flag'] }, ...base, { page: 'Product', items: ['Plum', 'Pear'] }, { page: 'Flag', items: ['FALSE'] }]),
    c('page-number-text', [{ pages: ['Product'] }, ...base, { page: 'Product', items: ['7'] }]),
    c('label-begins-lower', [...base, { filter: 'Region', type: 'CaptionBeginsWith', value1: 'e' }]),
    c('label-ends-wildcard', [...base, { filter: 'Region', type: 'CaptionEndsWith', value1: '?th' }]),
    c('label-not-between', [...base, { filter: 'Region', type: 'CaptionIsNotBetween', value1: 'D', value2: 'T' }]),
    c('label-greater-numbers', [{ rows: ['Code'] }, { value: { field: 'Qty' } }, { filter: 'Code', type: 'CaptionIsLessThan', value1: '25' }]),
    c('label-contains-long', [{ rows: ['Product'] }, { value: { field: 'Qty' } }, { filter: 'Product', type: 'CaptionContains', value1: 'honey' }]),
    c('value-between-average', [...base, { value: { field: 'Price', fn: 'average' } }, { filter: 'Region', type: 'ValueIsBetween', value1: 20, value2: 30, by: 'Average of Price' }]),
    c('value-less-count', [{ rows: ['Product'] }, { value: { field: 'Note', fn: 'count' } }, { filter: 'Product', type: 'ValueIsLessThan', value1: 8, by: 'Count of Note' }]),
    c('top-count-ties', [{ rows: ['Code'] }, { value: { field: 'Flag', fn: 'count' } }, { filter: 'Code', type: 'TopCount', value1: 2, by: 'Count of Flag' }]),
    c('bottom-percent-inner', [...two, { filter: 'Region', type: 'BottomPercent', value1: 40, by: 'Sum of Qty' }]),
    c('top-sum-errors', [{ rows: ['Region'] }, { value: { field: 'Amount' } }, { filter: 'Region', type: 'TopSum', value1: 500, by: 'Sum of Amount' }]),
    c('date-before', [{ rows: ['Date'] }, { value: { field: 'Qty' } }, { filter: 'Date', type: 'Before', value1: { d: serial(2024, 7, 1) } }]),
    c('date-after-equal', [{ rows: ['Date'] }, { value: { field: 'Qty' } }, { filter: 'Date', type: 'AfterOrEqualTo', value1: { d: serial(2025, 4, 22) } }]),
    c('date-quarter-4', [{ rows: ['Date'] }, { value: { field: 'Qty' } }, { filter: 'Date', type: 'AllDatesInPeriodQuarter4' }]),
    c('date-times', [{ rows: ['When'] }, { value: { field: 'Qty' } }, { filter: 'When', type: 'DateBetween', value1: { d: serial(2025, 1, 2) }, value2: { d: serial(2025, 2, 3) }, wholeDays: true }]),
    c('column-label-and-row-value', [{ rows: ['Region'] }, { columns: ['Product'] }, { value: { field: 'Qty' } }, { filter: 'Product', type: 'CaptionDoesNotContain', value1: 'p' }, { filter: 'Region', type: 'ValueIsGreaterThan', value1: 40, by: 'Sum of Qty' }]),
    c('hidden-page-and-top', [{ pages: ['Flag'] }, ...two, { page: 'Flag', items: ['TRUE'] }, { hide: 'Product', items: ['Plum'] }, { filter: 'Region', type: 'TopCount', value1: 1, by: 'Sum of Qty' }]),
    c('filter-then-cleared', [...base, { filter: 'Region', type: 'CaptionEquals', value1: 'South' }, { read: true }, { filter: 'Region', type: 'CaptionEquals', value1: 'North' }]),
  ]
}
