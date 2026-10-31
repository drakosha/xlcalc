/**
 * The set `checked-values` of fields dragged to the Values area of Excel's PivotTable Fields list (`checked.ts`): the
 * function Excel gives a field of numbers, with empty cells, with text among them, of text, of dates, of booleans, of
 * empty cells alone; and the Columns area again (`checked-columns` did not drag), a field of dates over years.
 */
import type { CheckedCase, CheckedSet } from './checked.js'

export const buildCheckedCases: CheckedSet['buildCheckedCases'] = (): CheckedCase[] => [
  ...['Num', 'NumBlank', 'NumText', 'Text', 'Years', 'Bool', 'Blank'].map((name): CheckedCase => ({ id: name, fields: [{ name, to: 'values' }] })),
  { id: 'dates-columns', fields: [{ name: 'Years', to: 'columns' }] },
  { id: 'month-columns', fields: [{ name: 'OneMonth', to: 'columns' }] },
]
