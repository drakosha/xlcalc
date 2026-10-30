/**
 * The set `show`: each of the 15 calculations of Show Values As over rows of Region and Product and columns of Flag:
 * shares of the total, the row, the column and the parents; the index; differences and percentages from an item, from
 * the previous and the next one; running totals and their percentage; ranks; over values with errors and empty items,
 * and beside a plain data field.
 */
import type { PivotCase, PivotOp, ValueField } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const grid: PivotOp[] = [{ rows: ['Region', 'Product'] }, { columns: ['Flag'] }]
  const one = (id: string, v: Omit<ValueField, 'field'>, field = 'Price', axes = grid): PivotCase => ({ id, ops: [...axes, { value: { field, ...v } }] })
  return [
    one('normal', { show: 'normal' }),
    one('total', { show: 'percentOfTotal' }),
    one('row', { show: 'percentOfRow' }),
    one('column', { show: 'percentOfCol' }),
    one('index', { show: 'index' }),
    one('parent-row', { show: 'percentOfParentRow' }),
    one('parent-column', { show: 'percentOfParentCol' }, 'Price', [{ rows: ['Region'] }, { columns: ['Flag', 'Code'] }]),
    one('parent-region', { show: 'percentOfParent', base: 'Region' }),
    one('parent-product', { show: 'percentOfParent', base: 'Product' }),
    one('run-region', { show: 'runTotal', base: 'Region' }),
    one('run-product', { show: 'runTotal', base: 'Product' }),
    one('run-flag', { show: 'runTotal', base: 'Flag' }),
    one('run-percent-region', { show: 'percentOfRunningTotal', base: 'Region' }),
    one('run-percent-product', { show: 'percentOfRunningTotal', base: 'Product' }),
    one('rank-up-product', { show: 'rankAscending', base: 'Product' }),
    one('rank-down-product', { show: 'rankDescending', base: 'Product' }),
    one('rank-down-region', { show: 'rankDescending', base: 'Region' }),
    one('difference-item', { show: 'difference', base: 'Region', item: 'East' }),
    one('difference-previous', { show: 'difference', base: 'Region', item: '(previous)' }),
    one('difference-next', { show: 'difference', base: 'Product', item: '(next)' }),
    one('difference-flag', { show: 'difference', base: 'Flag', item: 'FALSE' }),
    one('percent-diff-item', { show: 'percentDiff', base: 'Region', item: 'north' }),
    one('percent-diff-previous', { show: 'percentDiff', base: 'Product', item: '(previous)' }),
    one('percent-item', { show: 'percent', base: 'Region', item: 'West' }),
    one('percent-next', { show: 'percent', base: 'Region', item: '(next)' }),
    one('percent-blank', { show: 'percent', base: 'Region', item: '(blank)' }),
    one('errors-total', { show: 'percentOfTotal' }, 'Amount'),
    one('errors-run', { show: 'runTotal', base: 'Region' }, 'Amount'),
    one('errors-rank', { show: 'rankDescending', base: 'Region' }, 'Amount'),
    one('errors-difference', { show: 'difference', base: 'Region', item: '(previous)' }, 'Amount'),
    one('count-total', { fn: 'count', show: 'percentOfTotal' }, 'Note'),
    one('rows-only-index', { show: 'index' }, 'Price', [{ rows: ['Region'] }]),
    one('rows-only-column', { show: 'percentOfCol' }, 'Price', [{ rows: ['Region'] }]),
    { id: 'beside-plain', ops: [...grid, { value: { field: 'Price' } }, { value: { field: 'Price', caption: 'Share', show: 'percentOfParentRow' } }] },
    { id: 'tabular-parent', ops: [{ rows: ['Region', 'Product'] }, { value: { field: 'Qty', show: 'percentOfParent', base: 'Region' } }, { layout: 'tabular' }] },
    { id: 'sorted-run', ops: [{ rows: ['Region'] }, { value: { field: 'Price', show: 'runTotal', base: 'Region' } }, { sort: 'Region', order: 'descending' }] },
  ]
}
