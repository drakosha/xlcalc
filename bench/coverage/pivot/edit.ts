/**
 * The set `edit`: text and numbers put into the parts of a pivot table's area through `Value2` and typed through
 * `Formula2`: a row item's label, a column item's, a field's caption, a data field's, `Grand Total`, the header of the
 * compact form, a value, a subtotal, a page field's cell and its caption, an empty cell inside `TableRange2`; a name an
 * item or a field already has; a cell cleared. The answer is Excel's error, or the area after the edit and the book
 * with its definition.
 */
import type { EditValue, PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  // P!A1 the page field Product, A3 Sum of Qty and Column Labels, A4 Row Labels, FALSE, TRUE, Grand Total, A5:A9 East,
  // North, South, West, (blank), A10 Grand Total; in the tabular form A3 Region, B3 Code, A4 East, B4 10, A8 East Total
  const grid: PivotOp[] = [{ pages: ['Product'] }, { rows: ['Region'] }, { columns: ['Flag'] }, { value: { field: 'Qty' } }]
  const tabular: PivotOp[] = [{ rows: ['Region', 'Code'] }, { value: { field: 'Qty' } }, { layout: 'tabular' }]
  const at = (id: string, base: PivotOp[], cell: string, to: EditValue): PivotCase => ({ id, ops: [...base, { read: true }, { edit: `P!${cell}`, to }] })
  return [
    at('row-item', grid, 'A5', 'Eastern'),
    at('row-item-typed', grid, 'A5', { typed: 'Eastern' }),
    at('row-item-number', grid, 'A5', 7),
    at('row-item-same', grid, 'A5', 'East'),
    at('row-item-case', grid, 'A5', 'EAST'),
    at('row-item-taken', grid, 'A5', 'West'),
    at('row-item-taken-case', grid, 'A5', 'west'),
    at('row-item-other-field-item', grid, 'A5', 'Apple'),
    at('row-item-field-name', grid, 'A5', 'Region'),
    at('row-item-blank', grid, 'A9', 'none'),
    at('row-item-cleared', grid, 'A5', null),
    at('row-item-formula', grid, 'A5', { formula: '"E"&"ast2"' }),
    at('column-item', grid, 'B4', 'No'),
    at('column-item-taken', grid, 'B4', 'TRUE'),
    at('column-item-typed', grid, 'C4', { typed: 'yes' }),
    at('row-header', grid, 'A4', 'Regions'),
    at('column-header', grid, 'B3', 'Flags'),
    at('data-caption', grid, 'A3', 'Quantity'),
    at('data-caption-typed', grid, 'A3', { typed: 'Qty total' }),
    at('data-caption-source', grid, 'A3', 'Qty'),
    at('data-caption-other-field', grid, 'A3', 'Price'),
    at('data-caption-cleared', grid, 'A3', null),
    at('grand-total-row', grid, 'A10', 'All'),
    at('grand-total-column', grid, 'D4', 'Total'),
    at('value', grid, 'B5', 5),
    at('value-typed', grid, 'B5', { typed: '5' }),
    at('value-cleared', grid, 'B5', null),
    at('grand-value', grid, 'D10', 1),
    at('outside', grid, 'F5', 'x'),
    at('page-item', grid, 'B1', 'Apple'),
    at('page-item-typed', grid, 'B1', { typed: 'Banana' }),
    at('page-item-missing', grid, 'B1', 'Kiwi'),
    at('page-caption', grid, 'A1', 'Fruit'),
    at('empty-inside', grid, 'A2', 'x'),
    at('empty-header', grid, 'C3', 'x'),
    at('tabular-field', tabular, 'A3', 'Area'),
    at('tabular-field-taken', tabular, 'A3', 'Code'),
    at('tabular-inner-field', tabular, 'B3', 'Codes'),
    at('tabular-subtotal', tabular, 'A8', 'East all'),
    at('tabular-subtotal-value', tabular, 'C8', 1),
    at('tabular-empty', tabular, 'A5', 'x'),
    at('tabular-inner-item', tabular, 'B4', 'ten'),
    at('tabular-inner-item-taken', tabular, 'B4', '30'),
  ]
}
