/**
 * The held-out set `checked-held` of fields placed from Excel's PivotTable Fields list (`checked.ts`), over the rows of
 * `kinds2`: each column checked alone, and three fields in turn (numbers, dates, numbers with empty cells).
 */
import type { CheckedCase, CheckedSet } from './checked.js'
import { kindColumns2 } from './kinds-data.js'

export const buildCheckedCases: CheckedSet['buildCheckedCases'] = (): CheckedCase[] => [
  ...kindColumns2().map((c) => ({ id: c.name, data: 'kinds2', fields: [{ name: c.name }] })),
  { id: 'three', data: 'kinds2', fields: [{ name: 'Money' }, { name: 'TurnOfYear' }, { name: 'HalfBlank' }] },
  { id: 'two-dates', data: 'kinds2', fields: [{ name: 'TwoMonths' }, { name: 'YearsTimes' }] },
]
