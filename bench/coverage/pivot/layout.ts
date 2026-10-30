/**
 * The set `layout`: the compact, outline and tabular forms with two and three row fields; labels repeated; subtotals at
 * the top and bottom, none, several functions; grand totals of rows and columns on and off; a blank line after each
 * item; the table's own captions (`Grand Total`, `Row Labels`, `Column Labels`); a form per field; two column fields.
 */
import type { PivotCase, PivotOp } from './cases.js'

const LAYOUTS = ['compact', 'outline', 'tabular'] as const

export function buildPivotCases(): PivotCase[] {
  const two: PivotOp[] = [{ rows: ['Region', 'Product'] }, { value: { field: 'Price' } }]
  const three: PivotOp[] = [{ rows: ['Region', 'Flag', 'Code'] }, { value: { field: 'Qty' } }]
  const grid: PivotOp[] = [{ rows: ['Region'] }, { columns: ['Flag'] }, { value: { field: 'Price' } }]
  return [
    ...LAYOUTS.map((l): PivotCase => ({ id: `two-${l}`, ops: [...two, { layout: l }] })),
    ...LAYOUTS.map((l): PivotCase => ({ id: `three-${l}`, ops: [...three, { layout: l }] })),
    { id: 'repeat-outline', ops: [...three, { layout: 'outline' }, { repeatAll: true }] },
    { id: 'repeat-tabular', ops: [...three, { layout: 'tabular' }, { repeatAll: true }] },
    { id: 'repeat-field', ops: [...three, { layout: 'tabular' }, { field: 'Region', props: { repeatLabels: true } }] },
    { id: 'repeat-compact', ops: [...three, { repeatAll: true }] },
    ...LAYOUTS.map((l): PivotCase => ({ id: `bottom-${l}`, ops: [...three, { layout: l }, { field: 'Region', props: { subtotalsTop: false } }, { field: 'Flag', props: { subtotalsTop: false } }] })),
    { id: 'top-outline', ops: [...two, { layout: 'outline' }, { field: 'Region', props: { subtotalsTop: true } }] },
    ...LAYOUTS.map((l): PivotCase => ({ id: `no-subtotals-${l}`, ops: [...three, { layout: l }, { subtotals: 'Region', fns: 'none' }] })),
    { id: 'subtotals-three', ops: [...two, { layout: 'tabular' }, { subtotals: 'Region', fns: ['sum', 'count', 'average'] }] },
    { id: 'subtotals-three-compact', ops: [...two, { subtotals: 'Region', fns: ['max', 'min'] }] },
    { id: 'subtotals-inner', ops: [...three, { layout: 'outline' }, { subtotals: 'Flag', fns: ['count'] }] },
    { id: 'subtotals-columns', ops: [{ rows: ['Region'] }, { columns: ['Flag', 'Code'] }, { value: { field: 'Qty' } }, { subtotals: 'Flag', fns: ['sum', 'max'] }] },
    { id: 'no-row-grand', ops: [...grid, { table: { rowGrand: false } }] },
    { id: 'no-column-grand', ops: [...grid, { table: { columnGrand: false } }] },
    { id: 'no-grand', ops: [...grid, { table: { rowGrand: false, columnGrand: false } }] },
    { id: 'no-grand-rows-only', ops: [...two, { table: { columnGrand: false } }] },
    { id: 'blank-compact', ops: [...two, { field: 'Region', props: { blankLine: true } }] },
    { id: 'blank-tabular', ops: [...three, { layout: 'tabular' }, { field: 'Region', props: { blankLine: true } }, { field: 'Flag', props: { blankLine: true } }] },
    { id: 'captions', ops: [...grid, { table: { grandTotalName: 'All of it', rowHeader: 'Regions', columnHeader: 'Flags' } }] },
    { id: 'captions-tabular', ops: [...two, { layout: 'tabular' }, { table: { grandTotalName: 'Everything' } }, { field: 'Region', props: { caption: 'Area' } }] },
    { id: 'form-per-field', ops: [...three, { field: 'Region', props: { form: 'tabular' } }, { field: 'Flag', props: { form: 'outline' } }] },
    { id: 'two-columns', ops: [{ rows: ['Region'] }, { columns: ['Flag', 'Code'] }, { value: { field: 'Qty' } }] },
    { id: 'two-columns-tabular', ops: [{ rows: ['Region', 'Flag'] }, { columns: ['Code'] }, { value: { field: 'Qty' } }, { layout: 'tabular' }] },
    { id: 'two-values-tabular', ops: [{ rows: ['Region', 'Flag'] }, { value: { field: 'Price' } }, { value: { field: 'Qty', fn: 'average' } }, { layout: 'tabular' }] },
    { id: 'two-values-rows-outline', ops: [{ rows: ['Region', 'Flag'] }, { value: { field: 'Price' } }, { value: { field: 'Qty', fn: 'average' } }, { valuesOn: 'row' }, { layout: 'outline' }] },
    { id: 'page-tabular', ops: [{ pages: ['Flag', 'Code'] }, ...two, { layout: 'tabular' }] },
    { id: 'empty-text', ops: [{ rows: ['Region', 'Product'] }, { columns: ['Flag'] }, { value: { field: 'Price' } }, { layout: 'tabular' }, { table: { emptyText: '0' } }] },
    { id: 'phases', ops: [...two, { read: true }, { layout: 'tabular' }, { read: true }, { move: 'Product', to: 'column' }, { read: true }, { move: 'Region', to: 'hidden' }] },
  ]
}
