/** Held-out mixed-value merges and overlapping custom lists, predicted after the ordering rules were fixed. */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const sorts: PivotOp[] = [{ read: true }, { sort: 'Region', order: 'descending', by: 'Sum of Qty' }, { read: true }, { sort: 'Region', order: 'ascending', by: 'Sum of Qty' }]
  const cases: PivotCase[] = [600, 777].map(n => ({ id: `mixed-${n}`, data: `order-mixed-${n}`, ops: [{ rows: ['Region'] }, { value: { field: 'Qty' } }, ...sorts] }))
  const lists: (string | null)[][] = [
    ['May', 'June', 'Jun'], ['Jun', 'June', 'May'],
    ['Sat', 'Saturday', 'Dec', 'December'], ['December', 'Dec', 'Saturday', 'Sat'],
    ['February', 'Monday', 'Mon', 'Feb', 'mar'], ['mar', 'Feb', 'Mon', 'Monday', 'February'],
    ['Zebra', 'Tue', null, 'April', 'May', 'Apr', 'Tuesday'], ['Tuesday', 'Apr', 'May', 'April', null, 'Tue', 'Zebra'],
  ]
  lists.forEach((labels, k) => {
    const edits: PivotOp[] = Array.from({ length: 48 }, (_, i) => ({ edit: `D!I${i + 2}`, to: labels[i % labels.length] }))
    cases.push({ id: `lists-${k}`, ops: [...edits, { cache: { missingItems: 'none' } }, { refresh: true }, { rows: ['Note'] }, { value: { field: 'Qty' } }, { read: true }, { sort: 'Note', order: 'descending' }, { read: true }, { sort: 'Note', order: 'ascending' }, { read: true }, { refresh: true }] })
  })
  return cases
}
