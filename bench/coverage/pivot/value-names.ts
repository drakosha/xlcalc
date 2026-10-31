/**
 * The set `value-names`, run in en-US and ru-RU: what a value field's name becomes when its function changes
 * (`DataField.Function`), and what a calculated field added through the object model does on Values. Over Region on
 * rows and Qty on Values: the default name changed to Count and then to Average; a name of its own; a name of its own
 * that is another function's default name in en-US and in ru-RU, each language's default name of Sum given as a name;
 * the second field's default name (`Sum of Qty2`); a name set back to the default's text; a name and a function
 * changed together; the default name in lower case; a calculated field added (`CalculatedFields.Add`) and then its
 * formula changed on Values. Each step is a phase: the copy saved after naming is the book of `value-names-reopen`.
 */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const rows: PivotOp = { rows: ['Region'] }
  const qty = (caption?: string): PivotOp => ({ value: { field: 'Qty', ...(caption !== undefined ? { caption } : {}) } })
  const count: PivotOp = { setValue: 1, fn: 'count' }
  const c = (id: string, ops: PivotOp[]): PivotCase => ({ id, ops })
  return [
    c('default', [rows, qty(), { read: true }, count, { read: true }, { setValue: 1, fn: 'average' }]),
    c('own', [rows, qty('Quantity'), { read: true }, count]),
    c('other-default-en', [rows, qty('Average of Qty'), { read: true }, count]),
    c('other-default-ru', [rows, qty('Среднее по полю Qty'), { read: true }, count]),
    c('sum-text-en', [rows, qty('Sum of Qty'), { read: true }, count]),
    c('sum-text-ru', [rows, qty('Сумма по полю Qty'), { read: true }, count]),
    c('second', [rows, qty(), qty(), { read: true }, { setValue: 2, fn: 'count' }]),
    c('named-back', [rows, qty(), { setValue: 1, caption: 'Quantity' }, { setValue: 1, caption: 'Sum of Qty' }, { read: true }, count]),
    c('name-and-function', [rows, qty(), { read: true }, { setValue: 1, fn: 'count', caption: 'N' }]),
    c('lower-case', [rows, qty('sum of qty'), { read: true }, count]),
    c('calculated', [rows, qty(), { calcField: 'Unit', formula: '=Price/Qty' }, { read: true }, { value: { field: 'Unit' } }, { read: true }, { calcFormula: 'Unit', formula: '=Price*2' }]),
  ]
}
