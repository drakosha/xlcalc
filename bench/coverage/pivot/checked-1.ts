/**
 * The set `checked-1` of fields placed from Excel's PivotTable Fields list (`checked.ts`): each column of `kinds`
 * checked alone (where its box sends it, which function a value field takes, whether dates are grouped and by what);
 * two numeric fields (Σ Values), two fields of dates, dates after a number, and dates and a number sent to the columns.
 */
import type { CheckedCase, CheckedSet } from './checked.js'
import { kindColumns } from './kinds-data.js'

export const buildCheckedCases: CheckedSet['buildCheckedCases'] = (): CheckedCase[] => [
  ...kindColumns().map((c) => ({ id: c.name, fields: [{ name: c.name }] })),
  { id: 'two-numbers', fields: [{ name: 'Num' }, { name: 'NumBlank' }] },
  { id: 'two-dates', fields: [{ name: 'Years' }, { name: 'OneYear' }] },
  { id: 'number-then-dates', fields: [{ name: 'Num' }, { name: 'Years' }] },
  { id: 'dates-columns', fields: [{ name: 'Years', to: 'columns' }] },
  { id: 'text-then-dates-columns', fields: [{ name: 'Text' }, { name: 'OneYear', to: 'columns' }] },
  { id: 'number-columns', fields: [{ name: 'Num', to: 'columns' }] },
]
