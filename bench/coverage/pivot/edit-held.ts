/**
 * The held-out set `edit-held`, over the rows of `stock`: the edits of `edit` in the outline form, with two data fields,
 * on items with numbers, dates and the empty item, and two edits in a row.
 */
import type { EditValue, PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  // P!A3 Region, B3 Sum of Qty, C3 Sum of Price, A4 Central... in the outline form
  const outline: PivotOp[] = [{ rows: ['Region'] }, { value: { field: 'Qty' } }, { value: { field: 'Price' } }, { layout: 'outline' }]
  // P!A3 Row Labels, B3 Sum of Qty, A4 the first code
  const codes: PivotOp[] = [{ rows: ['Code'] }, { value: { field: 'Qty' } }]
  const dates: PivotOp[] = [{ rows: ['Date'] }, { value: { field: 'Qty' } }]
  const at = (id: string, base: PivotOp[], edits: [cell: string, to: EditValue][]): PivotCase => ({ id, data: 'stock', ops: [...base, { read: true }, ...edits.map(([cell, to]): PivotOp => ({ edit: `P!${cell}`, to }))] })
  return [
    at('outline-field', outline, [['A3', 'Where']]),
    at('outline-item', outline, [['A4', 'Middle']]),
    at('outline-item-taken', outline, [['A4', 'South']]),
    at('outline-data-caption', outline, [['B3', 'Units']]),
    at('outline-data-caption-taken', outline, [['C3', 'Sum of Qty']]),
    at('outline-data-caption-other-source', outline, [['C3', 'Amount']]),
    at('outline-grand-total', outline, [['A10', 'Everything']]),
    at('outline-value', outline, [['C5', 1]]),
    at('code-item-number', codes, [['A4', 50]]),
    at('code-item-taken-number', codes, [['A4', 25]]),
    at('code-item-taken-text', codes, [['A4', '15']]),
    at('date-item', dates, [['A4', 'first']]),
    at('date-item-date', dates, [['A4', 45000]]),
    at('blank-item', [{ rows: ['Region'] }, { value: { field: 'Qty' } }], [['A9', '(none)']]),
    at('two-edits', outline, [['A4', 'Middle'], ['A5', 'Middle']]),
    at('rename-back', outline, [['A4', 'Middle'], ['A4', 'Central']]),
    at('header-then-item', [{ rows: ['Region'] }, { value: { field: 'Qty' } }], [['A3', 'Places'], ['A4', 'Mid']]),
    at('typed-value', outline, [['B4', { typed: '=1+1' }]]),
  ]
}
