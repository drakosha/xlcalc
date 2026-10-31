/**
 * The set `com-dates`: a field of dates placed on the rows or the columns through the object model, by `Orientation`,
 * `AddFields` and the field's `CubeField`, to see whether Excel 365 groups it by periods (time grouping) as its field
 * list does.
 */
import type { PivotCase } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  return [
    { id: 'orientation-columns', ops: [{ columns: ['Date'] }, { value: { field: 'Qty' } }] },
    { id: 'add-fields-rows', ops: [{ addFields: { rows: ['Date'] } }, { value: { field: 'Qty' } }] },
    { id: 'add-fields-columns', ops: [{ addFields: { columns: ['Date'] } }, { value: { field: 'Qty' } }] },
    { id: 'add-fields-times', ops: [{ addFields: { rows: ['When'] } }, { value: { field: 'Qty' } }] },
    { id: 'cube-field-rows', ops: [{ cubeField: 'Date', to: 'row' }, { value: { field: 'Qty' } }] },
  ]
}
