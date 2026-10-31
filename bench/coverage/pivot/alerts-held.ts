/**
 * The held-out set `alerts-held`, run once the engine's prediction for it was committed after `alerts` and
 * `alerts-clear` were fitted: over a pivot table in tabular form of the rows of `stock` (`P!A1:B2` two report filters,
 * `A4:C31` the table, `A4:C4` its fields' names, `A5:B31` its labels, `A9` a subtotal's, `C5:C31` its values), a range
 * copied and pasted with it selected (Ctrl+C, Ctrl+V: `Range.Copy`, `Worksheet.Paste`) onto its values, labels and
 * filters and out of it; entries into a selection (`Formula2` over a range) whose first cell is outside it, partly
 * outside it, over its fields' names, a subtotal and its filters; a formula over a value and over a field's name; a
 * label and a filter cleared; Fill Down out of it and into it; cuts onto its filters and of its Filters area; columns
 * inserted right of it; the blank row under its filters deleted; its table's rows deleted without its filters and with
 * them; a merge, a table and pastes beside it.
 */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const base: PivotOp[] = [{ pages: ['Product', 'Flag'] }, { rows: ['Region', 'Code'] }, { value: { field: 'Qty' } }, { layout: 'tabular' }]
  const watch = ['P!A1:F12']
  const at = (id: string, op: PivotOp): PivotCase => ({ id, data: 'stock', watch, ops: [...base, { read: true }, op] })
  return [
    at('ctrl-v-onto-value', { copyPaste: 'D!A2', at: 'C5' }),
    at('ctrl-v-onto-label', { copyPaste: 'D!A2', at: 'A5' }),
    at('ctrl-v-onto-filter', { copyPaste: 'D!A2', at: 'B1' }),
    at('ctrl-v-beside', { copyPaste: 'D!A2', at: 'F5' }),
    at('ctrl-v-out-of-it', { copyPaste: 'P!A5:C6', at: 'E7' }),
    at('typed-first-outside', { edit: 'P!A3:B5', to: { typed: 'x' } }),
    at('typed-partly-outside', { edit: 'P!B5:D5', to: { typed: '7' } }),
    at('typed-field-names', { edit: 'P!A4:B4', to: { typed: 'Area' } }),
    at('typed-subtotal', { edit: 'P!A9:A10', to: { typed: 'All ?' } }),
    at('typed-filters', { edit: 'P!B1:B2', to: { typed: 'TRUE' } }),
    at('formula-on-value', { edit: 'P!C5', to: { typed: '=1' } }),
    at('formula-on-field', { edit: 'P!A4', to: { formula: '"R"' } }),
    at('cleared-label', { edit: 'P!A5', to: null }),
    at('cleared-filter', { edit: 'P!B1', to: null }),
    at('fill-down-out-of-it', { struct: 'fillDown', to: 'C31:C33' }),
    at('fill-down-into-it', { struct: 'fillDown', to: 'C3:C5' }),
    at('cut-onto-filter', { struct: 'moveRange', sheet: 'D', range: 'A2', to: 'B2' }),
    at('cut-filters', { struct: 'moveRange', range: 'A1:B2', to: 'F1' }),
    at('insert-columns-right', { struct: 'insertColumns', range: 'D1:E1' }),
    at('delete-blank-row', { struct: 'deleteRows', range: 'A3' }),
    at('delete-table-rows', { struct: 'deleteRows', range: 'A4:A31' }),
    at('delete-all-rows', { struct: 'deleteRows', range: 'A1:A31' }),
    at('merge-beside', { struct: 'mergeCells', range: 'D5:E5' }),
    at('table-beside', { struct: 'addTable', range: 'E5:F6', target: 'Beside' }),
    at('paste-blank-row', { paste: 'q', at: 'A3' }),
    at('paste-from-blank-row', { paste: 'a\tb', at: 'C3' }),
  ]
}
