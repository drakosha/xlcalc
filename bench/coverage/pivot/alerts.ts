/**
 * The set `alerts`: Excel's alert for each change of a pivot table's cells it refuses, as COM gives it with the alerts
 * off (the error's text in place of the window), for `reason` of a refusal with the code `pivot`; and what Excel does
 * with text entered into several of its cells at once. Over the pivot table of the set `edit` (`P!A1` the page field
 * Product, `A3:D10` the table, its rows' labels `A5:A9`, its values `B5:D10`): rows and columns inserted and deleted
 * through it and through the blank row under its Filters area; a cell copied (`Range.Copy` with a destination) and cut
 * (`Range.Cut`) onto its values, its labels and its Filters area, part of it cut away and all of it moved; Fill Down,
 * a merge and a table over it; text set through `Value2` and typed through `Formula2` into several of its cells at once
 * (as Ctrl+Enter enters a selection), and cleared; text of several lines pasted from the clipboard over its labels and
 * values (`Set-Clipboard`, `Worksheet.Paste`).
 */
import type { PivotCase, PivotOp } from './cases.js'

export function buildPivotCases(): PivotCase[] {
  const grid: PivotOp[] = [{ pages: ['Product'] }, { rows: ['Region'] }, { columns: ['Flag'] }, { value: { field: 'Qty' } }]
  const at = (id: string, op: PivotOp): PivotCase => ({ id, ops: [...grid, { read: true }, op] })
  return [
    at('insert-row-inside', { struct: 'insertRows', range: 'A6' }),
    at('delete-row-inside', { struct: 'deleteRows', range: 'A6' }),
    at('insert-column-inside', { struct: 'insertColumns', range: 'C1' }),
    at('delete-column-inside', { struct: 'deleteColumns', range: 'C1' }),
    at('insert-row-under-filters', { struct: 'insertRows', range: 'A2' }),
    at('insert-row-above', { struct: 'insertRows', range: 'A1' }),
    at('copy-onto-value', { struct: 'copyRange', sheet: 'D', range: 'A2', to: 'B6' }),
    at('copy-onto-label', { struct: 'copyRange', sheet: 'D', range: 'A2', to: 'A6' }),
    at('copy-onto-filter', { struct: 'copyRange', sheet: 'D', range: 'A2', to: 'B1' }),
    at('copy-over-edge', { struct: 'copyRange', sheet: 'D', range: 'A2:B3', to: 'D9' }),
    at('copy-out-of-it', { struct: 'copyRange', range: 'A5:B6', to: 'G5' }),
    at('cut-onto-value', { struct: 'moveRange', sheet: 'D', range: 'A2', to: 'B6' }),
    at('cut-part-of-it', { struct: 'moveRange', range: 'A5:B6', to: 'G5' }),
    at('cut-all-of-it', { struct: 'moveRange', range: 'A1:D10', to: 'G1' }),
    at('fill-down', { struct: 'fillDown', to: 'B5:B7' }),
    at('merge', { struct: 'mergeCells', range: 'B6:C6' }),
    at('table', { struct: 'addTable', range: 'A5:B6', target: 'Over' }),
    at('typed-labels', { edit: 'P!A5:A6', to: { typed: 'Eastern' } }),
    at('typed-labels-same', { edit: 'P!A5:A6', to: { typed: 'East' } }),
    at('typed-label-and-value', { edit: 'P!A5:B5', to: { typed: 'x' } }),
    at('typed-values', { edit: 'P!B5:B6', to: { typed: '5' } }),
    at('typed-formula-labels', { edit: 'P!A5:A6', to: { typed: '=1' } }),
    at('set-labels', { edit: 'P!A5:A6', to: 'Eastern' }),
    at('set-values', { edit: 'P!B5:B6', to: 5 }),
    at('cleared-labels', { edit: 'P!A5:A6', to: null }),
    at('cleared-values', { edit: 'P!B5:B6', to: null }),
    at('cleared-all', { edit: 'P!A1:D10', to: null }),
    at('paste-label', { paste: 'Eastern', at: 'A5' }),
    at('paste-labels', { paste: 'Eastern\nNorthern', at: 'A5' }),
    at('paste-labels-same', { paste: 'East\nNorth', at: 'A5' }),
    at('paste-labels-swapped', { paste: 'North\nEast', at: 'A5' }),
    at('paste-values', { paste: '1\n2', at: 'B5' }),
    at('paste-label-and-value', { paste: 'x\t1', at: 'A5' }),
    at('paste-header', { paste: 'Regions', at: 'A4' }),
  ]
}
