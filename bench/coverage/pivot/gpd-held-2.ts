/**
 * The held-out set `gpd-held-2`: GETPIVOTDATA where the outer row field's items are collapsed (some, all; tabular and
 * compact), with the second argument just outside the pivot table (one and two rows above it, a row below, a column to
 * the right; one data field, several on the columns with the Values row shown and hidden), and with Σ Values outside a
 * numeric column field.
 */
import type { PivotCase, PivotOp } from './cases.js'
import { gpd, gpdOver } from './gpd.js'

export function buildPivotCases(): PivotCase[] {
  const p = '"Sum of Price"'
  const two: PivotOp[] = [{ rows: ['Region', 'Product'] }, { value: { field: 'Price' } }]
  const outer = [gpd(p), gpd(p, '"Region"', '"East"'), gpd(p, '"Region"', '"South"'), gpd(p, '"Region"', '"East"', '"Product"', '"Apple"'), gpd(p, '"Region"', '"South"', '"Product"', '"Apple"'), gpd(p, '"Product"', '"Apple"')]
  const at = (cell: string, data = p) => `GETPIVOTDATA(${data},P!${cell})`
  const twoValues: PivotOp[] = [{ rows: ['Region'] }, { value: { field: 'Price' } }, { value: { field: 'Qty', fn: 'count' } }, { layout: 'outline' }]
  return [
    gpdOver('collapsed-some-tabular', [...two, { layout: 'tabular' }, { collapse: 'Region', items: ['East', 'North'] }], outer),
    gpdOver('collapsed-all-tabular', [...two, { layout: 'tabular' }, { collapse: 'Region', items: 'all' }], outer),
    gpdOver('collapsed-some-compact', [...two, { collapse: 'Region', items: ['East'] }], outer),
    gpdOver('collapsed-all-compact', [...two, { collapse: 'Region', items: 'all' }], outer),
    gpdOver('collapsed-columns', [{ rows: ['Region'] }, { columns: ['Flag', 'Code'] }, { value: { field: 'Qty' } }, { collapse: 'Flag', items: ['TRUE'] }], [gpd('"Sum of Qty"', '"Flag"', 'TRUE'), gpd('"Sum of Qty"', '"Flag"', 'TRUE', '"Code"', '10'), gpd('"Sum of Qty"', '"Region"', '"East"', '"Flag"', 'TRUE'), gpd('"Sum of Qty"', '"Flag"', 'FALSE', '"Code"', '20')]),
    gpdOver('around-one-value', [{ rows: ['Region'] }, { value: { field: 'Price' } }], [at('$A$2'), at('$A$1'), at('$B$2'), at('$A$10'), at('$B$10'), at('$C$3'), at('$C$9'), at('$A$3'), at('$B$9')]),
    gpdOver('around-values-row-shown', [...twoValues, { columns: ['Flag'] }], [at('$A$2'), at('$A$1'), at('$A$11', '"Count of Qty"'), at('$H$5', '"Count of Qty"'), at('$B$4', '"Count of Qty"')]),
    gpdOver('around-values-row-hidden', [...twoValues, { table: { valuesRow: false } }], [at('$A$2'), at('$A$1'), at('$A$10', '"Count of Qty"'), at('$D$5', '"Count of Qty"'), at('$B$3', '"Count of Qty"')]),
    gpdOver('around-values-row-on', [...twoValues, { table: { valuesRow: true } }], [at('$A$2'), at('$A$1'), at('$A$11', '"Count of Qty"'), at('$D$5', '"Count of Qty"'), at('$B$3', '"Count of Qty"')]),
    gpdOver('values-outside-numbers', [{ rows: ['Region', 'Product'] }, { columns: ['Code'] }, { value: { field: 'Price' } }, { value: { field: 'Qty' } }, { valuesOn: 'column', position: 1 }], [gpd(p, '"Region"', '"East"'), gpd('"Sum of Qty"', '"Region"', '"East"'), gpd(p, '"Region"', '"East"', '"Code"', '10'), gpd('"Sum of Qty"', '"Code"', '30'), gpd(p, '"Code"', '"20"'), gpd(p)]),
    gpdOver('values-outside-numbers-tabular', [{ rows: ['Region', 'Product'] }, { columns: ['Code'] }, { value: { field: 'Price' } }, { value: { field: 'Qty' } }, { valuesOn: 'column', position: 1 }, { layout: 'tabular' }], [gpd(p, '"Region"', '"West"'), gpd('"Sum of Qty"', '"Region"', '"West"', '"Code"', '20'), gpd(p)]),
  ]
}
