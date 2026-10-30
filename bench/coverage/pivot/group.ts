/**
 * The set `group`: dates grouped by years, quarters, months, days (and a number of days), several levels at once, with a
 * start and an end; dates and times by hours and minutes; numbers by intervals, with Excel's own start and end and with
 * its own, values outside them; text items grouped by hand (two groups, the empty item among them); a group on the
 * columns; ungrouping; and fields Excel refuses to group (text among the dates, numbers with text).
 */
import { serial } from './data.js'
import type { Period, PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const dates: PivotOp[] = [{ rows: ['Date'] }, { value: { field: 'Qty' } }]
  const when: PivotOp[] = [{ rows: ['When'] }, { value: { field: 'Qty' } }]
  const by = (id: string, periods: Period[], base = dates, field = 'Date'): PivotCase => ({ id, ops: [...base, { group: field, periods }] })
  return [
    by('years', ['years']),
    by('quarters', ['quarters']),
    by('months', ['months']),
    by('days', ['days']),
    by('months-years', ['months', 'years']),
    by('quarters-years', ['quarters', 'years']),
    by('days-months-years', ['days', 'months', 'quarters', 'years']),
    { id: 'days-7', ops: [...dates, { group: 'Date', periods: ['days'], by: 7 }] },
    { id: 'days-30-start', ops: [...dates, { group: 'Date', periods: ['days'], by: 30, start: { d: serial(2024, 1, 1) } }] },
    { id: 'months-start-end', ops: [...dates, { group: 'Date', periods: ['months', 'years'], start: { d: serial(2024, 3, 1) }, end: { d: serial(2025, 6, 30) } }] },
    by('hours', ['hours'], when, 'When'),
    by('minutes', ['minutes'], when, 'When'),
    by('hours-days', ['hours', 'days', 'months'], when, 'When'),
    by('seconds', ['seconds'], when, 'When'),
    { id: 'qty-2', ops: [{ rows: ['Qty'] }, { value: { field: 'Price' } }, { group: 'Qty', by: 2 }] },
    { id: 'qty-3-own', ops: [{ rows: ['Qty'] }, { value: { field: 'Price' } }, { group: 'Qty', start: 2, end: 7, by: 3 }] },
    { id: 'price-10', ops: [{ rows: ['Price'] }, { value: { field: 'Qty' } }, { group: 'Price', by: 10 }] },
    { id: 'price-10-numbers', ops: [{ edit: 'D!E14', to: 3 }, { refresh: true }, { rows: ['Price'] }, { value: { field: 'Qty' } }, { group: 'Price', by: 10 }] },
    { id: 'code-15', ops: [{ rows: ['Code'] }, { value: { field: 'Qty' } }, { group: 'Code', by: 15 }] },
    { id: 'amount-text', ops: [{ rows: ['Amount'] }, { value: { field: 'Qty' } }, { group: 'Amount', by: 100 }] },
    { id: 'text-dates', ops: [{ edit: 'D!C2', to: 'soon' }, { refresh: true }, ...dates, { group: 'Date', periods: ['months'] }] },
    { id: 'items', ops: [{ rows: ['Region'] }, { value: { field: 'Price' } }, { groupItems: 'Region', items: ['East', 'West'] }] },
    { id: 'items-two', ops: [{ rows: ['Region'] }, { value: { field: 'Price' } }, { groupItems: 'Region', items: ['East', 'West'] }, { groupItems: 'Region', items: ['North', 'South'] }] },
    { id: 'items-blank', ops: [{ rows: ['Region'] }, { value: { field: 'Price' } }, { groupItems: 'Region', items: ['(blank)', 'North'] }] },
    { id: 'items-columns', ops: [{ rows: ['Region'] }, { columns: ['Product'] }, { value: { field: 'Price' } }, { groupItems: 'Product', items: ['Apple', 'Banana'] }] },
    { id: 'items-tabular', ops: [{ rows: ['Region'] }, { value: { field: 'Price' } }, { layout: 'tabular' }, { groupItems: 'Region', items: ['East', 'South'] }] },
    { id: 'columns-months', ops: [{ rows: ['Region'] }, { columns: ['Date'] }, { value: { field: 'Qty' } }, { group: 'Date', periods: ['quarters', 'years'] }] },
    { id: 'ungroup', ops: [...dates, { group: 'Date', periods: ['months', 'years'] }, { read: true }, { ungroup: 'Date' }] },
    { id: 'refresh-new-date', ops: [...dates, { group: 'Date', periods: ['months', 'years'] }, { read: true }, { edit: 'D!C2', to: serial(2026, 2, 3) }, { refresh: true }] },
  ]
}
