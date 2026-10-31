/** Held-out value sorts: equal totals, new item counts, and ties within a parent. */
import type { PivotCase, PivotOp } from './cases.js'
import { ORDER_COUNTS } from './order-data.js'

export function buildPivotCases(): PivotCase[] {
  const sorts: PivotOp[] = [{ read: true }, { sort: 'Region', order: 'descending', by: 'Sum of Qty' }, { read: true }, { sort: 'Region', order: 'ascending', by: 'Sum of Qty' }]
  return [
    ...ORDER_COUNTS.map((n): PivotCase => ({ id: `items-${n}`, data: `order-${n}`, ops: [{ rows: ['Region'] }, { value: { field: 'Qty' } }, ...sorts] })),
    ...[7, 17, 33].map((n): PivotCase => ({ id: `inner-${n}`, data: `order-${n}`, ops: [{ rows: ['Product', 'Region'] }, { value: { field: 'Qty' } }, ...sorts] })),
  ]
}
