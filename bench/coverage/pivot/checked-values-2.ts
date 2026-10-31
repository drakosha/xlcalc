/**
 * The set `checked-values-2`: the fields of `checked-values` the mouse did not reach, out of the list's view, dragged
 * again with the field scrolled into view: to Values (text, dates, booleans, errors, empty cells alone, numbers with a
 * date) and to Columns (dates over years, dates within a month).
 */
import type { CheckedCase, CheckedSet } from './checked.js'

export const buildCheckedCases: CheckedSet['buildCheckedCases'] = (): CheckedCase[] => [
  ...['Text', 'Years', 'Bool', 'Err', 'Blank', 'NumDate'].map((name): CheckedCase => ({ id: name, fields: [{ name, to: 'values' }] })),
  { id: 'dates-columns', fields: [{ name: 'Years', to: 'columns' }] },
  { id: 'month-columns', fields: [{ name: 'OneMonth', to: 'columns' }] },
]
