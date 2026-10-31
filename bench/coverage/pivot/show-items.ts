/**
 * The set `show-items`: which items Show Values As takes. Shares, running totals, differences from the previous item,
 * ranks and the index with items hidden on rows and columns, a page field's item or items, Top 10 and a label filter,
 * items with no data shown; a base item hidden, a base field on no axis or on the Filters area; items sorted by the
 * values a calculation shows; value fields on the rows; other functions than the sum.
 */
import type { PivotCase, PivotOp, ValueField } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const grid: PivotOp[] = [{ rows: ['Region', 'Product'] }, { columns: ['Flag'] }]
  const region: PivotOp[] = [{ rows: ['Region'] }]
  const one = (id: string, axes: PivotOp[], v: Omit<ValueField, 'field'>, more: PivotOp[] = [], field = 'Price'): PivotCase => ({ id, ops: [...axes, { value: { field, ...v } }, ...more] })
  return [
    one('hidden-total', grid, { show: 'percentOfTotal' }, [{ hide: 'Region', items: ['West'] }]),
    one('hidden-column', grid, { show: 'percentOfRow' }, [{ hide: 'Flag', items: ['FALSE'] }]),
    one('hidden-index', grid, { show: 'index' }, [{ hide: 'Region', items: ['South'] }]),
    one('hidden-run', region, { show: 'runTotal', base: 'Region' }, [{ hide: 'Region', items: ['North'] }]),
    one('hidden-previous', grid, { show: 'difference', base: 'Region', item: '(previous)' }, [{ hide: 'Region', items: ['North'] }]),
    one('hidden-next', grid, { show: 'percent', base: 'Product', item: '(next)' }, [{ hide: 'Product', items: ['Banana'] }]),
    one('hidden-base', region, { show: 'percent', base: 'Region', item: 'North' }, [{ hide: 'Region', items: ['North'] }]),
    one('hidden-rank', region, { show: 'rankDescending', base: 'Region' }, [{ hide: 'Region', items: ['East'] }]),
    one('hidden-parent', grid, { show: 'percentOfParent', base: 'Region' }, [{ hide: 'Product', items: ['Apple'] }]),
    one('page-total', [...region, { pages: ['Product'] }], { show: 'percentOfTotal' }, [{ page: 'Product', items: ['Apple'] }]),
    one('page-run', [...region, { pages: ['Product'] }], { show: 'runTotal', base: 'Region' }, [{ page: 'Product', items: ['Apple', 'Cherry'] }]),
    {
      id: 'top-share',
      ops: [
        ...region,
        { value: { field: 'Price' } },
        { value: { field: 'Price', caption: 'Share', show: 'percentOfTotal' } },
        { value: { field: 'Price', caption: 'Run', show: 'runTotal', base: 'Region' } },
        { filter: 'Region', type: 'TopCount', value1: 2, by: 'Sum of Price' },
      ],
    },
    one('label-parent-row', grid, { show: 'percentOfParentRow' }, [{ filter: 'Product', type: 'CaptionDoesNotEqual', value1: 'Apple' }]),
    one('empty-previous', grid, { show: 'difference', base: 'Region', item: '(previous)' }, [{ field: 'Product', props: { showAllItems: true } }]),
    one('empty-percent', grid, { show: 'percent', base: 'Product', item: 'Banana' }, [{ field: 'Product', props: { showAllItems: true } }]),
    one('base-off-axes', region, { show: 'difference', base: 'Flag', item: 'TRUE' }),
    one('base-page', [...region, { pages: ['Flag'] }], { show: 'percent', base: 'Flag', item: 'TRUE' }),
    one('sort-by-rank', region, { caption: 'Rank', show: 'rankDescending', base: 'Region' }, [{ sort: 'Region', order: 'ascending', by: 'Rank' }]),
    one('sort-by-difference', region, { caption: 'Diff', show: 'difference', base: 'Region', item: 'East' }, [{ sort: 'Region', order: 'descending', by: 'Diff' }]),
    {
      id: 'values-on-rows',
      ops: [
        { rows: ['Region'] },
        { columns: ['Flag'] },
        { value: { field: 'Price', show: 'percentOfCol' } },
        { value: { field: 'Qty', caption: 'Run Qty', show: 'runTotal', base: 'Flag' } },
        { valuesOn: 'row' },
      ],
    },
    one('average-previous', grid, { fn: 'average', show: 'percentDiff', base: 'Region', item: '(previous)' }),
    one('count-rank', region, { fn: 'count', show: 'rankAscending', base: 'Region' }, [], 'Note'),
  ]
}
