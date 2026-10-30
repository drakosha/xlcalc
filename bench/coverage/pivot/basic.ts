/**
 * The set `basic`: one row field and one data field with each of the 11 functions, over a field with empty cells,
 * text, a boolean and errors (Amount) and over one without errors (Price); a column field; two data fields with their
 * Σ Values on the columns and on the rows; row items of numbers, dates, booleans, long text and text in two cases; the
 * three kinds of source; a data field's caption and format; what errors and empty cells show; a page field.
 */
import { AGGS, type PivotCase } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const byRegion = (field: string, fn: (typeof AGGS)[number]): PivotCase['ops'] => [{ rows: ['Region'] }, { value: { field, fn } }]
  return [
    ...AGGS.map((fn): PivotCase => ({ id: `amount-${fn}`, ops: byRegion('Amount', fn) })),
    ...AGGS.map((fn): PivotCase => ({ id: `price-${fn}`, ops: byRegion('Price', fn) })),
    { id: 'region-count', ops: byRegion('Region', 'count') },
    { id: 'region-count-nums', ops: byRegion('Region', 'countNums') },
    { id: 'note-count', ops: [{ rows: ['Product'] }, { value: { field: 'Note', fn: 'count' } }] },
    { id: 'columns', ops: [{ rows: ['Region'] }, { columns: ['Product'] }, { value: { field: 'Price' } }] },
    { id: 'columns-only', ops: [{ columns: ['Region'] }, { value: { field: 'Price' } }] },
    { id: 'values-only', ops: [{ value: { field: 'Price' } }] },
    { id: 'rows-only', ops: [{ rows: ['Region'] }] },
    { id: 'two-values', ops: [{ rows: ['Region'] }, { value: { field: 'Price' } }, { value: { field: 'Qty', fn: 'count' } }] },
    { id: 'two-values-rows', ops: [{ rows: ['Region'] }, { value: { field: 'Price' } }, { value: { field: 'Qty', fn: 'count' } }, { valuesOn: 'row' }] },
    { id: 'two-values-rows-first', ops: [{ rows: ['Region'] }, { value: { field: 'Price' } }, { value: { field: 'Qty', fn: 'average' } }, { valuesOn: 'row', position: 1 }] },
    { id: 'two-values-columns', ops: [{ rows: ['Region'] }, { columns: ['Flag'] }, { value: { field: 'Price' } }, { value: { field: 'Qty', fn: 'max' } }] },
    { id: 'two-values-alone', ops: [{ value: { field: 'Price' } }, { value: { field: 'Amount', fn: 'count' } }] },
    { id: 'same-field-twice', ops: [{ rows: ['Region'] }, { value: { field: 'Price', fn: 'sum' } }, { value: { field: 'Price', fn: 'average' } }] },
    { id: 'code-rows', ops: [{ rows: ['Code'] }, { value: { field: 'Qty' } }] },
    { id: 'date-rows', ops: [{ rows: ['Date'] }, { value: { field: 'Qty' } }] },
    { id: 'flag-rows', ops: [{ rows: ['Flag'] }, { value: { field: 'Qty' } }] },
    { id: 'product-rows', ops: [{ rows: ['Product'] }, { value: { field: 'Qty' } }] },
    { id: 'note-rows', ops: [{ rows: ['Note'] }, { value: { field: 'Qty' } }] },
    { id: 'amount-rows', ops: [{ rows: ['Amount'] }, { value: { field: 'Qty', fn: 'count' } }] },
    { id: 'caption-format', ops: [{ rows: ['Region'] }, { value: { field: 'Price', caption: 'Total price', format: '#,##0.00' } }] },
    { id: 'shown-errors', ops: [{ rows: ['Region'] }, { columns: ['Product'] }, { value: { field: 'Amount' } }, { table: { errorText: 'err', emptyText: '-' } }] },
    { id: 'source-table', source: 'table', ops: byRegion('Price', 'sum') },
    { id: 'source-name', source: 'name', ops: byRegion('Price', 'sum') },
    { id: 'two-rows', ops: [{ rows: ['Region', 'Product'] }, { value: { field: 'Price' } }] },
    { id: 'page', ops: [{ pages: ['Product'] }, { rows: ['Region'] }, { value: { field: 'Price' } }] },
    { id: 'page-one', ops: [{ pages: ['Product'] }, { rows: ['Region'] }, { value: { field: 'Price' } }, { page: 'Product', items: ['Apple'] }] },
  ]
}
