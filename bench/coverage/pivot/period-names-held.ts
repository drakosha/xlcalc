/**
 * The held-out set `period-names-held`, run in en-US and ru-RU after the engine named groupings in the workbook's
 * language, over the rows of `stock`: quarters and years from a start of their own, days and months on the columns,
 * minutes and hours, days by 14, years to an end of their own, whole numbers by a fraction from a start of their own.
 */
import { serial } from './data.js'
import type { Period, PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const qty: PivotOp = { value: { field: 'Qty' } }
  const c = (id: string, ops: PivotOp[]): PivotCase => ({ id, data: 'stock', ops })
  const rows = (field: string, periods: Period[], more: Partial<Extract<PivotOp, { group: string }>> = {}): PivotOp[] => [{ rows: [field] }, qty, { layout: 'tabular' }, { group: field, periods, ...more }]
  return [
    c('quarters-years-start', rows('Date', ['quarters', 'years'], { start: { d: serial(2024, 6, 1) } })),
    c('days-months-columns', [{ rows: ['Region'] }, { columns: ['Date'] }, qty, { group: 'Date', periods: ['days', 'months'] }]),
    c('minutes-hours', rows('When', ['minutes', 'hours'])),
    c('days-14', rows('Date', ['days'], { by: 14 })),
    c('years-end', rows('Date', ['years'], { end: { d: serial(2025, 3, 1) } })),
    // Excel's object model takes Qty off the rows as it becomes a value field: nothing is grouped (the next case groups it)
    c('qty-fraction-start', [{ rows: ['Qty'] }, { value: { field: 'Qty', fn: 'count' } }, { group: 'Qty', start: 2, end: 9, by: 1.5 }]),
    c('qty-fraction-price', [{ rows: ['Qty'] }, { value: { field: 'Price', fn: 'count' } }, { group: 'Qty', start: 2, end: 9, by: 1.5 }]),
  ]
}
