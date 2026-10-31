/**
 * The set `com-values`: a field put on Values without a function through the object model (`AddDataField` with the
 * field alone), over the rows of `kinds`: the function Excel gives a field of numbers, of numbers with empty cells,
 * text or dates among them, of text, of dates, of booleans, of errors, of empty cells alone, of numbers kept as text.
 */
import type { PivotCase } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  return ['Num', 'NumBlank', 'NumText', 'NumDate', 'Text', 'Years', 'Bool', 'Err', 'Blank', 'NumAsText'].map((f) => ({ id: f, data: 'kinds', ops: [{ valueDefault: f }] }))
}
