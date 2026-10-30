/** The names Excel gives a PivotTable's parts: a value field's (`Sum of Amount`), a new PivotTable's (`PivotTable1`). */
import type { PivotFunction } from '../contract.js'

/** The function's word in a value field's name, as Excel for Windows en-US names it (`countNums` is `Count` too). */
const WORD: Record<PivotFunction, string> = { sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count', stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp' }

/** A value field's name by its function and its source field's name; en-US whatever the workbook's language. */
export const valueFieldName = (fn: PivotFunction, field: string) => `${WORD[fn]} of ${field}`
