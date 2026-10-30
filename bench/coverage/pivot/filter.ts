/**
 * The set `filter`: hidden items; a page field with one item, several, all, the empty one; label filters (equals,
 * begins with, ends with, contains, comparisons, between, with wildcards); value filters; top 10 by items, percent and
 * sum; date filters; and filters together (a label and a value filter on one field, hidden items under a value
 * filter, a filter on an inner field, on a column field, a page field with a filter on the rows).
 */
import { serial } from './data.js'
import type { FilterType, FilterValue, PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const base: PivotOp[] = [{ rows: ['Region'] }, { value: { field: 'Price' } }]
  const two: PivotOp[] = [{ rows: ['Region', 'Product'] }, { value: { field: 'Price' } }]
  const dates: PivotOp[] = [{ rows: ['Date'] }, { value: { field: 'Qty' } }]
  const label = (id: string, type: FilterType, value1: FilterValue, value2?: FilterValue): PivotCase => ({ id, ops: [...base, { filter: 'Region', type, value1, ...(value2 !== undefined ? { value2 } : {}) }] })
  const value = (id: string, type: FilterType, value1: number, value2?: number): PivotCase => ({ id, ops: [...base, { filter: 'Region', type, by: 'Sum of Price', value1, ...(value2 !== undefined ? { value2 } : {}) }] })
  const date = (id: string, type: FilterType, value1?: FilterValue, value2?: FilterValue): PivotCase => ({ id, ops: [...dates, { filter: 'Date', type, ...(value1 !== undefined ? { value1 } : {}), ...(value2 !== undefined ? { value2 } : {}) }] })
  return [
    { id: 'hide-one', ops: [...base, { hide: 'Region', items: ['East'] }] },
    { id: 'hide-blank', ops: [...base, { hide: 'Region', items: ['(blank)'] }] },
    { id: 'hide-three', ops: [...base, { hide: 'Region', items: ['East', 'West', 'South'] }] },
    { id: 'hide-inner', ops: [...two, { hide: 'Product', items: ['Apple', 'Cherry'] }] },
    { id: 'hide-column', ops: [{ rows: ['Region'] }, { columns: ['Code'] }, { value: { field: 'Qty' } }, { hide: 'Code', items: ['30'] }] },
    // two items are named 20 (a number and a text): COM finds neither by the name
    { id: 'hide-two-named', ops: [{ rows: ['Region'] }, { columns: ['Code'] }, { value: { field: 'Qty' } }, { hide: 'Code', items: ['20'] }] },
    { id: 'page-one', ops: [{ pages: ['Region'] }, { rows: ['Product'] }, { value: { field: 'Price' } }, { page: 'Region', items: ['West'] }] },
    { id: 'page-several', ops: [{ pages: ['Region'] }, { rows: ['Product'] }, { value: { field: 'Price' } }, { page: 'Region', items: ['West', 'East'] }] },
    { id: 'page-blank', ops: [{ pages: ['Region'] }, { rows: ['Product'] }, { value: { field: 'Price' } }, { page: 'Region', items: ['(blank)'] }] },
    { id: 'page-all', ops: [{ pages: ['Region'] }, { rows: ['Product'] }, { value: { field: 'Price' } }, { page: 'Region', items: ['West'] }, { read: true }, { page: 'Region', items: '(All)' }] },
    { id: 'page-two', ops: [{ pages: ['Region', 'Flag'] }, { rows: ['Product'] }, { value: { field: 'Price' } }, { page: 'Region', items: ['North'] }, { page: 'Flag', items: ['TRUE'] }] },
    label('equals', 'CaptionEquals', 'North'),
    label('equals-case', 'CaptionEquals', 'NORTH'),
    label('not-equal', 'CaptionDoesNotEqual', 'North'),
    label('begins', 'CaptionBeginsWith', 'N'),
    label('not-begins', 'CaptionDoesNotBeginWith', 's'),
    label('ends', 'CaptionEndsWith', 'th'),
    label('not-ends', 'CaptionDoesNotEndWith', 'st'),
    label('contains', 'CaptionContains', 'or'),
    label('not-contains', 'CaptionDoesNotContain', 'e'),
    label('greater', 'CaptionIsGreaterThan', 'N'),
    label('greater-equal', 'CaptionIsGreaterThanOrEqualTo', 'North'),
    label('less', 'CaptionIsLessThan', 'S'),
    label('less-equal', 'CaptionIsLessThanOrEqualTo', 'South'),
    label('between', 'CaptionIsBetween', 'E', 'S'),
    label('not-between', 'CaptionIsNotBetween', 'E', 'S'),
    label('wildcard-star', 'CaptionEquals', 'N*'),
    label('wildcard-mark', 'CaptionEquals', '?est'),
    label('blank-label', 'CaptionEquals', '(blank)'),
    { id: 'label-numbers', ops: [{ rows: ['Code'] }, { value: { field: 'Qty' } }, { filter: 'Code', type: 'CaptionIsGreaterThan', value1: '15' }] },
    { id: 'label-numbers-equal', ops: [{ rows: ['Code'] }, { value: { field: 'Qty' } }, { filter: 'Code', type: 'CaptionEquals', value1: '20' }] },
    value('value-greater', 'ValueIsGreaterThan', 250),
    value('value-greater-equal', 'ValueIsGreaterThanOrEqualTo', 301.91),
    value('value-less', 'ValueIsLessThan', 250),
    value('value-less-equal', 'ValueIsLessThanOrEqualTo', 237.8),
    value('value-equals', 'ValueEquals', 348.4),
    value('value-not-equal', 'ValueDoesNotEqual', 348.4),
    value('value-between', 'ValueIsBetween', 200, 320),
    value('value-not-between', 'ValueIsNotBetween', 200, 320),
    value('top-count', 'TopCount', 2),
    value('bottom-count', 'BottomCount', 2),
    value('top-percent', 'TopPercent', 50),
    value('bottom-percent', 'BottomPercent', 30),
    value('top-sum', 'TopSum', 600),
    value('bottom-sum', 'BottomSum', 300),
    { id: 'top-inner', ops: [...two, { filter: 'Product', type: 'TopCount', value1: 1, by: 'Sum of Price' }] },
    { id: 'top-errors', ops: [{ rows: ['Region'] }, { value: { field: 'Amount' } }, { filter: 'Region', type: 'TopCount', value1: 2, by: 'Sum of Amount' }] },
    date('date-after', 'After', { d: serial(2025, 1, 1) }),
    date('date-before-equal', 'BeforeOrEqualTo', { d: serial(2024, 5, 1) }),
    date('date-between', 'DateBetween', { d: serial(2024, 6, 1) }, { d: serial(2025, 3, 31) }),
    date('date-not-between', 'DateNotBetween', { d: serial(2024, 6, 1) }, { d: serial(2025, 3, 31) }),
    date('date-specific', 'SpecificDate', { d: serial(2024, 1, 1) }),
    date('date-january', 'AllDatesInPeriodJanuary'),
    date('date-quarter-2', 'AllDatesInPeriodQuarter2'),
    { id: 'label-and-value', ops: [...base, { table: { multipleFilters: true } }, { filter: 'Region', type: 'CaptionDoesNotEqual', value1: 'West' }, { filter: 'Region', type: 'TopCount', value1: 2, by: 'Sum of Price' }] },
    { id: 'label-then-value', ops: [...base, { filter: 'Region', type: 'CaptionDoesNotEqual', value1: 'West' }, { filter: 'Region', type: 'TopCount', value1: 2, by: 'Sum of Price' }] },
    { id: 'hidden-and-value', ops: [...base, { hide: 'Region', items: ['South'] }, { filter: 'Region', type: 'ValueIsGreaterThan', value1: 200, by: 'Sum of Price' }] },
    { id: 'inner-label', ops: [...two, { filter: 'Product', type: 'CaptionBeginsWith', value1: 'B' }] },
    { id: 'outer-value-inner-label', ops: [...two, { filter: 'Region', type: 'ValueIsGreaterThan', value1: 250, by: 'Sum of Price' }, { filter: 'Product', type: 'CaptionContains', value1: 'an' }] },
    { id: 'column-filter', ops: [{ rows: ['Region'] }, { columns: ['Product'] }, { value: { field: 'Price' } }, { filter: 'Product', type: 'TopCount', value1: 2, by: 'Sum of Price' }] },
    { id: 'page-and-label', ops: [{ pages: ['Flag'] }, ...base, { page: 'Flag', items: ['FALSE'] }, { filter: 'Region', type: 'CaptionContains', value1: 'th' }] },
  ]
}
