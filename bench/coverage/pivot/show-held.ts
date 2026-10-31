/**
 * The held-out set `show-held`, over the rows of `stock`: the calculations of Show Values As on other fields and
 * bases (rows of Product and Region, columns of Code), with grouped dates, with two data fields, in the tabular form.
 */
import type { PivotCase, PivotOp, ValueField } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const grid: PivotOp[] = [{ rows: ['Product', 'Region'] }, { columns: ['Flag'] }]
  const c = (id: string, ops: PivotOp[]): PivotCase => ({ id, data: 'stock', ops })
  const one = (id: string, v: Omit<ValueField, 'field'>, field = 'Qty', axes = grid): PivotCase => c(id, [...axes, { value: { field, ...v } }])
  const months: PivotOp[] = [{ rows: ['Date'] }, { group: 'Date', periods: ['months', 'years'] }]
  return [
    one('total', { show: 'percentOfTotal' }, 'Price'),
    one('row', { show: 'percentOfRow' }),
    one('column-codes', { show: 'percentOfCol' }, 'Qty', [{ rows: ['Region'] }, { columns: ['Code'] }]),
    one('index-codes', { show: 'index' }, 'Price', [{ rows: ['Region'] }, { columns: ['Code'] }]),
    one('parent-row', { show: 'percentOfParentRow' }, 'Price'),
    one('parent-column', { show: 'percentOfParentCol' }, 'Qty', [{ rows: ['Product'] }, { columns: ['Region', 'Flag'] }]),
    one('parent-product', { show: 'percentOfParent', base: 'Product' }),
    one('run-region', { show: 'runTotal', base: 'Region' }),
    one('run-percent-product', { show: 'percentOfRunningTotal', base: 'Product' }, 'Price'),
    one('rank-up-region', { show: 'rankAscending', base: 'Region' }),
    one('rank-down-codes', { show: 'rankDescending', base: 'Code' }, 'Price', [{ rows: ['Code'] }, { columns: ['Flag'] }]),
    one('difference-item', { show: 'difference', base: 'Product', item: 'Pear' }),
    one('difference-next', { show: 'difference', base: 'Region', item: '(next)' }),
    one('percent-diff-flag', { show: 'percentDiff', base: 'Flag', item: 'TRUE' }),
    one('percent-previous', { show: 'percent', base: 'Product', item: '(previous)' }, 'Price'),
    c('months-difference', [...months, { value: { field: 'Qty', show: 'difference', base: 'Date', item: '(previous)' } }]),
    c('months-run', [...months, { value: { field: 'Price', show: 'runTotal', base: 'Date' } }]),
    c('two-values', [...grid, { value: { field: 'Qty' } }, { value: { field: 'Qty', caption: 'Share', show: 'percentOfParent', base: 'Product' } }, { value: { field: 'Price', caption: 'Rank', show: 'rankDescending', base: 'Region' } }]),
    c('tabular-index', [...grid, { value: { field: 'Price', show: 'index' } }, { layout: 'tabular' }]),
    one('errors-parent', { show: 'percentOfParentRow' }, 'Amount'),
  ]
}
