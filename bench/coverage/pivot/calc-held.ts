/**
 * The held-out set `calc-held`, over the rows of `stock`: calculated fields with functions of the sheet and errors
 * (IF over errors, IFERROR, powers, ABS, MIN, a division by zero on some items, booleans), on rows and columns with
 * subtotals; calculated items with subtotals and grand totals; a field and an item together; over two pivot tables of
 * a cache.
 */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const c = (id: string, ops: PivotOp[]): PivotCase => ({ id, data: 'stock', ops })
  const field = (id: string, formula: string, rows: string[] = ['Region'], more: PivotOp[] = []): PivotCase => c(id, [{ rows }, { calcField: 'Calc field', formula }, { value: { field: 'Calc field' } }, ...more])
  return [
    field('if-errors', '=IF(Amount>0,Amount,0)'),
    field('iferror-ratio', '=IFERROR(Price/Qty,0)', ['Region', 'Product']),
    field('power', '=Price^2/100'),
    field('abs', '=ABS(Amount)'),
    field('min', '=MIN(Price,10*Qty)', ['Product']),
    field('div-some', '=Price/(Qty-Qty+Code-15)', ['Code']),
    field('boolean', '=TRUE+Qty'),
    field('columns-subtotals', '=Price/Qty', ['Region', 'Flag'], [{ columns: ['Product'] }, { layout: 'outline' }]),
    field('ratio-tabular', '=Amount/Qty', ['Product', 'Region'], [{ layout: 'tabular' }]),
    c('item-subtotals', [{ rows: ['Region', 'Product'] }, { value: { field: 'Qty' } }, { calcItem: 'Two fruits', field: 'Product', formula: '=Plum+Pear' }]),
    c('item-grand', [{ rows: ['Region'] }, { columns: ['Flag'] }, { value: { field: 'Qty' } }, { calcItem: 'Middle less West', field: 'Region', formula: '=Central-West' }]),
    c('item-number', [{ rows: ['Code'] }, { value: { field: 'Price' } }, { calcItem: 'Five twice', field: 'Code', formula: "='5'*2" }]),
    c('item-lower', [{ rows: ['Region'] }, { value: { field: 'Qty' } }, { calcItem: 'East and west', field: 'Region', formula: '=east+West' }]),
    c('item-two-fields', [{ rows: ['Region', 'Product'] }, { value: { field: 'Qty' } }, { calcItem: 'Both', field: 'Region', formula: '=North+South' }, { calcItem: 'Pair', field: 'Product', formula: '=Fig+Apple' }]),
    c('field-and-item', [{ rows: ['Product'] }, { calcField: 'Unit', formula: '=Price/Qty' }, { value: { field: 'Unit' } }, { value: { field: 'Qty' } }, { calcItem: 'Plum and fig', field: 'Product', formula: '=Plum+Fig' }]),
    c('item-then-field', [{ rows: ['Region'] }, { value: { field: 'Price' } }, { calcItem: 'Twice', field: 'Region', formula: '=South*2' }, { calcField: 'Half', formula: '=Price/2' }, { value: { field: 'Half' } }]),
    c('shared-field', [{ rows: ['Region'] }, { create: 'PT2', at: 'P!E3' }, { rows: ['Product'], pt: 'PT2' }, { calcField: 'Unit', formula: '=Price/Qty' }, { value: { field: 'Unit' } }, { value: { field: 'Unit' }, pt: 'PT2' }]),
    c('shared-item', [{ rows: ['Region'] }, { value: { field: 'Qty' } }, { create: 'PT2', at: 'P!E3' }, { rows: ['Region'], pt: 'PT2' }, { value: { field: 'Price' }, pt: 'PT2' }, { calcItem: 'North and south', field: 'Region', formula: '=North+South', pt: 'PT2' }]),
  ]
}
