/**
 * The set `period-names`, run in en-US and ru-RU: the names Excel gives the fields a grouping of dates and times makes
 * (`Years (Date)`) and the labels of their items (quarters, months, days, hours, minutes, seconds, the `<` and `>`
 * groups of a start and an end), in the language of Excel's interface. Each period alone and all of them at once, in
 * the tabular layout, whose headers show the fields' names; a number of days; a start and an end of its own; numbers by
 * an interval with a fraction.
 */
import { serial } from './data.js'
import type { Period, PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const tabular: PivotOp = { layout: 'tabular' }
  const of = (field: string): PivotOp[] => [{ rows: [field] }, { value: { field: 'Qty' } }, tabular]
  const by = (id: string, field: string, periods: Period[], more: Partial<Extract<PivotOp, { group: string }>> = {}): PivotCase => ({ id, ops: [...of(field), { group: field, periods, ...more }] })
  return [
    by('years', 'Date', ['years']),
    by('quarters', 'Date', ['quarters']),
    by('months', 'Date', ['months']),
    by('days', 'Date', ['days']),
    by('hours', 'When', ['hours']),
    by('minutes', 'When', ['minutes']),
    by('seconds', 'When', ['seconds']),
    by('dates-all', 'Date', ['days', 'months', 'quarters', 'years']),
    by('times-all', 'When', ['seconds', 'minutes', 'hours', 'days', 'months', 'quarters', 'years']),
    by('days-7', 'Date', ['days'], { by: 7 }),
    by('start-end', 'Date', ['months', 'years'], { start: { d: serial(2024, 3, 1) }, end: { d: serial(2025, 6, 30) } }),
    // Price's one text made a number first, as the set group does
    { id: 'price-half', ops: [{ edit: 'D!E14', to: 3 }, { refresh: true }, ...of('Price'), { group: 'Price', by: 2.5 }] },
  ]
}
