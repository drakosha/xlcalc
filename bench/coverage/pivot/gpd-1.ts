/**
 * The set `gpd-1`: GETPIVOTDATA over the books of `basic`, in the forms books use: the data field by its caption and by
 * its source field's name, in another case, with spaces; items as text, numbers, dates, booleans, a reference and
 * built by TEXT; a subtotal, the grand total, part of the pairs; a page field; an item or a field that is not there;
 * the second argument at another cell of the area, outside it, a range; a ratio and a difference of two calls; and
 * formulas over a pivot table in another book.
 */
import type { PivotCase } from './cases.js'
import { gpd, gpdCase } from './gpd.js'

const LONG = '"Dragonfruit from the far southern orchards, packed in wooden crates of twelve and shipped by sea"'

export function buildPivotCases(): PivotCase[] {
  const p = '"Sum of Price"'
  const q = '"Sum of Qty"'
  return [
    gpdCase(
      'columns',
      'basic/columns',
      [
        gpd(p), gpd('"Price"'), gpd('"sum of price"'), gpd('"PRICE"'), gpd('" Sum of Price"'), gpd('"Sum of Price "'), gpd('"Count of Price"'),
        gpd(p, '"Region"', '"North"'), gpd(p, '"Region"', '"north"'), gpd(p, '"region"', '"NORTH"'), gpd(p, '"Region"', '"Nowhere"'), gpd(p, '"Region"', '"(blank)"'), gpd(p, '"Region"', '""'),
        gpd(p, '"Region"', '"North"', '"Product"', '"Apple"'), gpd(p, '"Product"', '"Apple"', '"Region"', '"North"'), gpd(p, '"Product"', '"Banana"'), gpd(p, '"Product"', '"100"'), gpd(p, '"Product"', '100'), gpd(p, '"Product"', LONG),
        gpd(p, '"Flag"', 'TRUE'), gpd(p, '"Region"'), gpd(p, '"Region"', 'G!$H$1'), 'GETPIVOTDATA(G!$H$2,P!$A$3)', gpd(p, '"Region"', '{"North","East"}'),
        'GETPIVOTDATA("Sum of Price",P!$C$7)', 'GETPIVOTDATA("Sum of Price",P!$A$1)', 'GETPIVOTDATA("Sum of Price",P!A3:B4)', 'GETPIVOTDATA("Sum of Price",P!$M$11)',
        `${gpd(p, '"Region"', '"West"')}/${gpd(p)}`, `${gpd(p, '"Region"', '"West"')}-${gpd(p, '"Region"', '"East"')}`,
      ],
      ['GETPIVOTDATA("Sum of Price",$A$3,"Region","South")', 'GETPIVOTDATA("Price",A5)'],
      ['West', 'Sum of Price'],
    ),
    gpdCase('two-rows', 'basic/two-rows', [
      gpd(p), gpd(p, '"Region"', '"North"'), gpd(p, '"Region"', '"North"', '"Product"', '"Apple"'), gpd(p, '"Product"', '"Apple"'), gpd(p, '"Product"', '"Apple"', '"Region"', '"(blank)"'),
      gpd(p, '"Region"', '"(blank)"'), gpd(p, '"Region"', '"East"', '"Product"', '"Nowhere"'), gpd(p, '"Region"', '"South"', '"Product"', LONG), 'GETPIVOTDATA("Sum of Price",P!$A$10)',
    ]),
    gpdCase('code', 'basic/code-rows', [
      gpd(q), gpd(q, '"Code"', '10'), gpd(q, '"Code"', '"10"'), gpd(q, '"Code"', '20'), gpd(q, '"Code"', '"20"'), gpd(q, '"Code"', 'TEXT(30,"0")'), gpd(q, '"Code"', '10.0'),
      gpd(q, '"Code"', '"10.0"'), gpd(q, '"Code"', '1E1'), gpd(q, '"Code"', '40'), gpd(q, '"Code"', 'G!$H$1'), gpd(q, '"Code"', 'G!$H$2'),
    ], [], [30, '20']),
    gpdCase('date', 'basic/date-rows', [
      gpd(q, '"Date"', 'DATE(2024,1,1)'), gpd(q, '"Date"', '45292'), gpd(q, '"Date"', '"1/1/2024"'), gpd(q, '"Date"', '"2024-01-01"'), gpd(q, '"Date"', 'TEXT(45292,"m/d/yyyy")'),
      gpd(q, '"Date"', 'TEXT(45292,"yyyy-mm-dd")'), gpd(q, '"Date"', '"January 1, 2024"'), gpd(q, '"Date"', 'DATEVALUE("1/1/2024")'), gpd(q, '"Date"', 'G!$H$1'), gpd(q, '"Date"', '45292.5'),
      gpd(q, '"Date"', '"(blank)"'), gpd(q, '"Date"', 'DATE(2024,1,2)'),
    ], [], [45292]),
    gpdCase('flag', 'basic/flag-rows', [gpd(q, '"Flag"', 'TRUE'), gpd(q, '"Flag"', '"TRUE"'), gpd(q, '"Flag"', '"true"'), gpd(q, '"Flag"', '1'), gpd(q, '"Flag"', 'FALSE'), gpd(q, '"Flag"', '1=1')]),
    gpdCase('two-values', 'basic/two-values', [
      gpd(p, '"Region"', '"East"'), gpd('"Count of Qty"', '"Region"', '"East"'), gpd('"Qty"', '"Region"', '"East"'), gpd('"Count of Qty"'), gpd('"Price"', '"Region"', '"West"'), gpd('"Sum of Qty"'),
    ]),
    gpdCase('two-values-rows', 'basic/two-values-rows', [gpd(p, '"Region"', '"East"'), gpd('"Count of Qty"', '"Region"', '"East"'), gpd('"Count of Qty"'), gpd(p)]),
    gpdCase('page-all', 'basic/page', [gpd(p), gpd(p, '"Product"', '"Apple"'), gpd(p, '"Product"', '"(All)"'), gpd(p, '"Region"', '"East"'), 'GETPIVOTDATA("Sum of Price",P!$A$1)', 'GETPIVOTDATA("Sum of Price",P!$B$1)']),
    gpdCase('page-one', 'basic/page-one', [gpd(p), gpd(p, '"Product"', '"Apple"'), gpd(p, '"Product"', '"Banana"'), gpd(p, '"Region"', '"East"'), gpd(p, '"Region"', '"East"', '"Product"', '"Apple"')]),
    gpdCase('errors', 'basic/amount-sum', [gpd('"Sum of Amount"'), gpd('"Sum of Amount"', '"Region"', '"North"'), gpd('"Sum of Amount"', '"Region"', '"East"'), gpd('"Sum of Amount"', '"Region"', '"South"'), `IFERROR(${gpd('"Amount"', '"Region"', '"North"')},0)`]),
    gpdCase('caption', 'basic/caption-format', [gpd('"Total price"'), gpd('"Price"'), gpd(p), gpd('"total PRICE"', '"Region"', '"North"')]),
    {
      id: 'other-book',
      from: 'basic/columns',
      ops: [{ external: 'basic/two-rows', cells: { 'G!A1': '=GETPIVOTDATA("Sum of Price",[Src.xlsx]P!$A$3)', 'G!A2': '=GETPIVOTDATA("Sum of Price",[Src.xlsx]P!$A$3,"Region","North")', 'G!A3': '=GETPIVOTDATA("Price",[Src.xlsx]P!$A$3,"Region","North","Product","Apple")', 'G!A4': '=GETPIVOTDATA("Sum of Price",[Src.xlsx]P!$A$3,"Region","Nowhere")' } }],
      watch: ['G!A1:A4'],
    },
  ]
}
