/**
 * The set `checked-columns` of fields dragged to the Columns area of Excel's PivotTable Fields list (`checked.ts`):
 * dates over years and within a year (are they grouped as on the rows), and a number (does it go there).
 */
import type { CheckedCase, CheckedSet } from './checked.js'

export const buildCheckedCases: CheckedSet['buildCheckedCases'] = (): CheckedCase[] => [
  { id: 'dates-columns', fields: [{ name: 'Years', to: 'columns' }] },
  { id: 'text-then-dates-columns', fields: [{ name: 'Text' }, { name: 'OneYear', to: 'columns' }] },
  { id: 'month-columns', fields: [{ name: 'OneMonth', to: 'columns' }] },
  { id: 'number-columns', fields: [{ name: 'Num', to: 'columns' }] },
]
