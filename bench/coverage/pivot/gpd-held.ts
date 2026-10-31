/**
 * The held-out set `gpd-held`: GETPIVOTDATA over the books of `layout`, `filter` and `sort`: page fields with one item
 * and several, hidden and filtered items, several data fields, subtotals of their own functions, custom captions of
 * the grand total, items of numbers and dates, the forms of the corpus over other layouts.
 */
import type { PivotCase } from './cases.js'
import { gpd, gpdCase } from './gpd.js'

export function buildPivotCases(): PivotCase[] {
  const p = '"Sum of Price"'
  const q = '"Sum of Qty"'
  return [
    gpdCase('tabular', 'layout/two-tabular', [gpd(p), gpd(p, '"Region"', '"South"'), gpd(p, '"Region"', '"South"', '"Product"', '"Cherry"'), gpd('"price"', '"Product"', '"Banana"', '"Region"', '"west"'), 'GETPIVOTDATA("Sum of Price",P!$B$10)']),
    gpdCase('three', 'layout/three-outline', [gpd(q, '"Region"', '"East"', '"Flag"', 'FALSE'), gpd(q, '"Region"', '"East"', '"Flag"', 'FALSE', '"Code"', '20'), gpd(q, '"Region"', '"East"', '"Flag"', 'FALSE', '"Code"', '"20"'), gpd(q, '"Flag"', 'TRUE'), gpd(q, '"Region"', '"North"', '"Code"', '10')]),
    gpdCase('subtotals-functions', 'layout/subtotals-three', [gpd(p, '"Region"', '"East"'), `GETPIVOTDATA("Sum of Price",P!$A$3,"Region","East","Count")`, `GETPIVOTDATA("Sum of Price",P!$A$3,"Region[East;Average]")`, `GETPIVOTDATA("Sum of Price",P!$A$3,"Region[East;Sum]")`, gpd(p)]),
    gpdCase('no-subtotals', 'layout/no-subtotals-tabular', [gpd(q, '"Region"', '"West"'), gpd(q, '"Region"', '"West"', '"Flag"', 'TRUE'), gpd(q)]),
    gpdCase('no-grand', 'layout/no-grand', [gpd(p), gpd(p, '"Region"', '"East"'), gpd(p, '"Flag"', 'TRUE'), gpd(p, '"Region"', '"East"', '"Flag"', 'TRUE')]),
    gpdCase('captions', 'layout/captions', [gpd(p), gpd(p, '"Region"', '"South"'), gpd('"All of it"'), gpd(p, '"Flag"', '"All of it"')]),
    gpdCase('caption-field', 'layout/captions-tabular', [gpd(p, '"Region"', '"North"'), gpd(p, '"Area"', '"North"'), gpd('"Everything"')]),
    gpdCase('two-values-tabular', 'layout/two-values-tabular', [gpd('"Average of Qty"', '"Region"', '"East"'), gpd('"Average of Qty"', '"Region"', '"East"', '"Flag"', 'FALSE'), gpd('"Qty"'), gpd(p, '"Flag"', 'FALSE')]),
    gpdCase('pages-two', 'filter/page-two', [gpd(p), gpd(p, '"Region"', '"North"'), gpd(p, '"Region"', '"South"'), gpd(p, '"Flag"', 'TRUE', '"Product"', '"Apple"')]),
    gpdCase('page-several', 'filter/page-several', [gpd(p), gpd(p, '"Region"', '"West"'), gpd(p, '"Product"', '"Apple"')]),
    gpdCase('hidden', 'filter/hide-one', [gpd(p), gpd(p, '"Region"', '"East"'), gpd(p, '"Region"', '"West"')]),
    gpdCase('top', 'filter/top-count', [gpd(p), gpd(p, '"Region"', '"North"'), gpd(p, '"Region"', '"South"')]),
    gpdCase('dates-filtered', 'filter/date-between', [gpd('"Sum of Qty"'), gpd(q, '"Date"', 'DATE(2024,9,1)'), gpd(q, '"Date"', 'DATE(2024,1,1)'), gpd(q, '"Date"', '"9/1/2024"')]),
    gpdCase('sorted', 'sort/value-down', [gpd(p, '"Region"', '"South"'), 'GETPIVOTDATA("Sum of Price",P!$B$4)']),
    gpdCase('manual', 'sort/refresh-manual-new', [gpd(p, '"Region"', '"Central"'), gpd(p, '"Region"', '"East"'), gpd(p)]),
  ]
}
