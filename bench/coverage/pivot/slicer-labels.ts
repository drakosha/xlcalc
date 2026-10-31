/**
 * The set `slicer-labels`: the captions of a slicer's items (`SlicerItem.Caption`, with `Name` and `Value`) over the
 * rows of `labels`, a field of each kind on the rows: text with a blank, numbers in a currency, a percentage, General and
 * `#,##0`, dates in the short date and in a long one, times, dates with times, times without a date among dates with
 * times, booleans, errors, numbers and text mixed; times in a book of the 1904 date system; a field on no axis; a field with the PivotTable's own number format; dates grouped by months, by months and years, and
 * numbers grouped. Run in en-US and ru-RU. Predicted before Excel ran it: `slicer-labels.predicted.json`
 * (`slicer-ui-predict.ts`).
 */
import type { PivotCase, PivotOp } from './cases.js'

const ITEMS = ['Name', 'Caption', 'Value']

/** A slicer over `field` named `Slicer_<field>`, its items read. */
const sliced = (field: string, name = `Slicer_${field.replace(/\W/g, '_')}`): PivotOp[] => [{ slicer: field, name }, { slicerGet: name, items: ITEMS }]

const FIELDS = ['Text', 'Money', 'Pct', 'Plain', 'Thou', 'Day', 'Long', 'Time', 'When', 'Clock', 'Flag', 'Err', 'Mixed']

export function buildPivotCases(): PivotCase[] {
  const value: PivotOp = { value: { field: 'N' } }
  return [
    ...FIELDS.map((f): PivotCase => ({ id: f.toLowerCase(), data: 'labels', ops: [{ rows: [f] }, value, ...sliced(f)] })),
    { id: 'money-no-axis', data: 'labels', ops: [value, ...sliced('Money')] },
    { id: 'plain-own-format', data: 'labels', ops: [{ rows: ['Plain'] }, value, { props: { format: '0.00' }, field: 'Plain' }, ...sliced('Plain')] },
    { id: 'day-months', data: 'labels', ops: [{ rows: ['Day'] }, value, { group: 'Day', periods: ['months'] }, ...sliced('Day')] },
    { id: 'day-months-years', data: 'labels', ops: [{ rows: ['Long'] }, value, { group: 'Long', periods: ['months', 'years'] }, ...sliced('Months (Long)', 'Slicer_Months'), ...sliced('Years (Long)', 'Slicer_Years')] },
    { id: 'time-1904', data: 'labels', dateSystem: 1904, ops: [{ rows: ['Time'] }, value, ...sliced('Time')] },
    { id: 'clock-1904', data: 'labels', dateSystem: 1904, ops: [{ rows: ['Clock'] }, value, ...sliced('Clock')] },
    { id: 'time-no-axis', data: 'labels', ops: [value, ...sliced('Time')] },
    { id: 'n-numbers', data: 'labels', ops: [{ rows: ['N'] }, { value: { field: 'Pct' } }, { group: 'N', start: 1, end: 8, by: 3 }, ...sliced('N')] },
  ]
}
