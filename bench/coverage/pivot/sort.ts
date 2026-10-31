/**
 * The set `sort`: AutoSort by labels up and down (text, numbers beside numbers kept as text, dates, long text), by a
 * data field's values (an outer and an inner field, a column field, values with errors), items moved by hand, Excel's
 * custom lists (month names) on and off, and how a sort and a hand-made order hold through an edit of the source and a
 * refresh (values changed, a new item).
 */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const base: PivotOp[] = [{ rows: ['Region'] }, { value: { field: 'Price' } }]
  const two: PivotOp[] = [{ rows: ['Region', 'Product'] }, { value: { field: 'Price' } }]
  const by = (field: string, value: string, fn: 'sum' | 'count' = 'sum'): PivotOp[] => [{ rows: [field] }, { value: { field: value, fn } }]
  const months = ['Mar', 'Jan', 'December', 'Feb', 'mon', 'Sunday', 'Apr', 'jan', 'Wed']
  const monthEdits: PivotOp[] = months.map((m, i): PivotOp => ({ edit: `D!I${2 + i}`, to: m }))
  return [
    { id: 'region-up', ops: [...base, { sort: 'Region', order: 'ascending' }] },
    { id: 'region-down', ops: [...base, { sort: 'Region', order: 'descending' }] },
    { id: 'code-down', ops: [...by('Code', 'Qty'), { sort: 'Code', order: 'descending' }] },
    { id: 'date-down', ops: [...by('Date', 'Qty'), { sort: 'Date', order: 'descending' }] },
    { id: 'product-down', ops: [...by('Product', 'Qty'), { sort: 'Product', order: 'descending' }] },
    { id: 'flag-down', ops: [...by('Flag', 'Qty'), { sort: 'Flag', order: 'descending' }] },
    { id: 'note-up', ops: [...by('Note', 'Qty'), { sort: 'Note', order: 'ascending' }] },
    { id: 'amount-up', ops: [...by('Amount', 'Qty', 'count'), { sort: 'Amount', order: 'ascending' }] },
    { id: 'value-down', ops: [...base, { sort: 'Region', order: 'descending', by: 'Sum of Price' }] },
    { id: 'value-up', ops: [...base, { sort: 'Region', order: 'ascending', by: 'Sum of Price' }] },
    { id: 'value-errors', ops: [{ rows: ['Region'] }, { value: { field: 'Amount' } }, { sort: 'Region', order: 'descending', by: 'Sum of Amount' }] },
    { id: 'value-errors-up', ops: [{ rows: ['Region'] }, { value: { field: 'Amount' } }, { sort: 'Region', order: 'ascending', by: 'Sum of Amount' }] },
    { id: 'inner-value', ops: [...two, { sort: 'Product', order: 'descending', by: 'Sum of Price' }] },
    { id: 'outer-and-inner', ops: [...two, { sort: 'Region', order: 'descending', by: 'Sum of Price' }, { sort: 'Product', order: 'ascending', by: 'Sum of Price' }] },
    { id: 'column-value', ops: [{ rows: ['Region'] }, { columns: ['Product'] }, { value: { field: 'Price' } }, { sort: 'Product', order: 'descending', by: 'Sum of Price' }] },
    { id: 'count-ties', ops: [...by('Code', 'Qty', 'count'), { sort: 'Code', order: 'descending', by: 'Count of Qty' }] },
    { id: 'manual-first', ops: [...base, { item: 'West', of: 'Region', position: 1 }] },
    { id: 'manual-last', ops: [...base, { item: 'East', of: 'Region', position: 5 }] },
    { id: 'manual-blank', ops: [...base, { item: '(blank)', of: 'Region', position: 1 }] },
    { id: 'manual-then-field', ops: [...base, { item: 'South', of: 'Region', position: 2 }, { read: true }, { columns: ['Flag'] }] },
    { id: 'months', ops: [...monthEdits, { refresh: true }, ...by('Note', 'Qty')] },
    { id: 'months-down', ops: [...monthEdits, { refresh: true }, ...by('Note', 'Qty'), { sort: 'Note', order: 'descending' }] },
    { id: 'months-no-lists', ops: [{ table: { customLists: false } }, ...monthEdits, { refresh: true }, ...by('Note', 'Qty')] },
    { id: 'refresh-value-sort', ops: [...base, { sort: 'Region', order: 'descending', by: 'Sum of Price' }, { read: true }, { edit: 'D!E3', to: 900 }, { refresh: true }] },
    { id: 'refresh-label-sort-new', ops: [...base, { sort: 'Region', order: 'descending' }, { read: true }, { edit: 'D!A4', to: 'Central' }, { edit: 'D!A6', to: 'Zulu' }, { refresh: true }] },
    { id: 'refresh-auto-off', ops: [...base, { sort: 'Region', order: 'descending' }, { sort: 'Region', order: 'manual' }, { read: true }, { edit: 'D!A4', to: 'Central' }, { edit: 'D!A6', to: 'Zulu' }, { refresh: true }] },
    { id: 'refresh-value-auto-off', ops: [...base, { sort: 'Region', order: 'descending', by: 'Sum of Price' }, { sort: 'Region', order: 'manual', by: 'Sum of Price' }, { read: true }, { edit: 'D!E3', to: 900 }, { refresh: true }] },
    { id: 'refresh-unsorted-new', ops: [...base, { read: true }, { edit: 'D!A4', to: 'Central' }, { edit: 'D!A6', to: 'Alpha' }, { refresh: true }] },
    { id: 'refresh-manual-new', ops: [...base, { item: 'West', of: 'Region', position: 1 }, { read: true }, { edit: 'D!A4', to: 'Central' }, { refresh: true }] },
  ]
}
