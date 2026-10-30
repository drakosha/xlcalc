/**
 * The set `calc`: calculated fields (a ratio of sums, over an item with two rows, `Amount/Qty` with (100, 1) and (100, 9)
 * giving 20, with subtotals and grand totals; division by zero; functions of the sheet; a constant; text and boolean
 * fields; one calculated field over another; a data field over one with another function or Show Values As) and
 * calculated items (sums and differences of items, an item with a long name, numeric items, with subtotals, beside a
 * calculated field); and both over two pivot tables of one cache, with grouping, which acts on both.
 */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  // the item Calc: two rows, Amount 100 and 100, Qty 1 and 9
  const twoRows: PivotOp[] = [
    { edit: 'D!A2', to: 'Calc' }, { edit: 'D!D2', to: 100 }, { edit: 'D!F2', to: 1 },
    { edit: 'D!A3', to: 'Calc' }, { edit: 'D!D3', to: 100 }, { edit: 'D!F3', to: 9 },
    { refresh: true },
  ]
  const field = (id: string, formula: string, rows: string[] = ['Region'], more: PivotOp[] = []): PivotCase => ({
    id,
    ops: [...twoRows, { rows }, { calcField: 'Calc field', formula }, { value: { field: 'Calc field' } }, ...more],
  })
  const item = (id: string, field: string, name: string, formula: string, more: PivotOp[] = []): PivotCase => ({
    id,
    ops: [{ rows: [field] }, { value: { field: 'Price' } }, { calcItem: name, field, formula }, ...more],
  })
  return [
    field('ratio', '=Amount/Qty'),
    field('ratio-two-rows', '=Amount/Qty', ['Region', 'Flag']),
    field('ratio-price', '=Price/Qty', ['Region', 'Product']),
    field('ratio-tabular', '=Price/Qty', ['Region', 'Flag'], [{ layout: 'tabular' }, { subtotals: 'Region', fns: ['sum', 'average'] }]),
    field('ratio-columns', '=Price/Qty', ['Region'], [{ columns: ['Flag'] }]),
    field('div-zero', '=Price/0'),
    field('div-zero-difference', '=Price/(Qty-Qty)'),
    field('if', '=IF(Qty>40,Price,0)'),
    field('round', '=ROUND(Price*1.2,1)'),
    field('max', '=MAX(Price,Qty*10)'),
    field('constant', '=100'),
    field('text-field', '=Region+1'),
    field('boolean-field', '=Flag*Qty'),
    field('errors', '=Amount*2'),
    field('iserror', '=IFERROR(Amount,0)+Price'),
    field('and-or', '=AND(Price>100,Qty>20)*1'),
    { id: 'over-another', ops: [...twoRows, { rows: ['Region'] }, { calcField: 'Unit', formula: '=Price/Qty' }, { calcField: 'Unit2', formula: "=Unit*2+'Qty'" }, { value: { field: 'Unit' } }, { value: { field: 'Unit2' } }] },
    { id: 'name-with-space', ops: [{ rows: ['Region'] }, { calcField: 'Unit price', formula: '=Price/Qty' }, { calcField: 'Half', formula: "='Unit price'/2" }, { value: { field: 'Half' } }] },
    { id: 'count-of', ops: [{ rows: ['Region'] }, { calcField: 'Unit', formula: '=Price/Qty' }, { value: { field: 'Unit', fn: 'count' } }] },
    { id: 'percent-of-total', ops: [{ rows: ['Region'] }, { calcField: 'Unit', formula: '=Price/Qty' }, { value: { field: 'Unit', show: 'percentOfTotal' } }] },
    { id: 'beside-sum', ops: [...twoRows, { rows: ['Region'] }, { value: { field: 'Amount' } }, { value: { field: 'Qty' } }, { calcField: 'Unit', formula: '=Amount/Qty' }, { value: { field: 'Unit' } }] },
    item('item-sum', 'Region', 'North and South', '=North+South'),
    item('item-difference', 'Region', 'North less South', '=North-South'),
    item('item-times', 'Region', 'Double West', '=West*2'),
    item('item-long', 'Product', 'Fruit', "=Apple+'Dragonfruit from the far southern orchards, packed in wooden crates of twelve and shipped by sea'"),
    item('item-numbers', 'Code', 'Ten and thirty', "='10'+'30'"),
    item('item-subtotals', 'Product', 'Two', '=Apple+Banana', [{ rows: ['Region'] }, { move: 'Region', to: 'row', position: 1 }]),
    item('item-columns', 'Region', 'Both', '=East+West', [{ columns: ['Flag'] }]),
    item('item-average', 'Region', 'Both', '=East+West', [{ value: { field: 'Qty', fn: 'average' } }]),
    item('item-div-zero', 'Region', 'Ratio', '=East/0'),
    { id: 'item-and-field', ops: [{ rows: ['Region'] }, { calcField: 'Unit', formula: '=Price/Qty' }, { value: { field: 'Unit' } }, { calcItem: 'Both', field: 'Region', formula: '=East+West' }] },
    {
      id: 'shared-cache',
      ops: [
        ...twoRows,
        { rows: ['Region'] },
        { create: 'PT2', at: 'P!E3' },
        { rows: ['Region'], pt: 'PT2' },
        { value: { field: 'Qty' }, pt: 'PT2' },
        { calcField: 'Unit', formula: '=Amount/Qty' },
        { value: { field: 'Unit' } },
        { read: true },
        { value: { field: 'Unit' }, pt: 'PT2' },
        { read: true },
        { calcItem: 'Both', field: 'Region', formula: '=East+West' },
      ],
    },
    {
      id: 'shared-group',
      ops: [
        { rows: ['Date'] },
        { value: { field: 'Qty' } },
        { create: 'PT2', at: 'P!E3' },
        { rows: ['Date'], pt: 'PT2' },
        { value: { field: 'Price' }, pt: 'PT2' },
        { read: true },
        { group: 'Date', periods: ['months', 'years'] },
        { read: true },
        { ungroup: 'Date', pt: 'PT2' },
      ],
    },
  ]
}
