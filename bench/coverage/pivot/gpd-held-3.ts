/**
 * The held-out set `gpd-held-3`: GETPIVOTDATA over items renamed in the PivotTable (a caption of their own): a row item
 * and a column item renamed by an entry (the books of `edit`), a row item renamed to a number's text, and items of an
 * outer field collapsed and then renamed (the books of `gpd-held-2`, tabular and compact). Each item is named by its
 * value, its caption, either in another case, a value of no item; beside an item not renamed.
 */
import type { PivotCase, PivotOp } from './cases.js'
import { gpd, gpdCase } from './gpd.js'

/** A case over a saved book whose item at `cell` is renamed to `to` before its formulas are typed. */
function renamed(id: string, from: string, cell: string, to: string, onG: string[]): PivotCase {
  const c = gpdCase(id, from, onG)
  return { ...c, ops: [{ edit: `P!${cell}`, to } as PivotOp, ...c.ops] }
}

export function buildPivotCases(): PivotCase[] {
  const q = '"Sum of Qty"'
  const p = '"Sum of Price"'
  // edit/row-item: East renamed Eastern on the rows (Region), Flag on the columns, Product on the Filters area
  // edit/column-item: FALSE of Flag renamed No on the columns
  // edit/row-item-number: East renamed 7
  const collapsed = [gpd(p, '"Region"', '"East"'), gpd(p, '"Region"', '"other"'), gpd(p, '"Region"', '"OTHER"'), gpd(p, '"Region"', '"east"'), gpd(p, '"Region"', '"North"'), gpd(p, '"Region"', '"South"')]
  return [
    gpdCase('row-renamed', 'edit/row-item', [gpd(q, '"Region"', '"East"'), gpd(q, '"Region"', '"Eastern"'), gpd(q, '"Region"', '"eastern"'), gpd(q, '"Region"', '"EAST"'), gpd(q, '"Region"', '"East"', '"Flag"', 'TRUE'), gpd(q, '"Region"', '"Eastern"', '"Flag"', 'FALSE'), gpd(q, '"Region"', '"North"'), gpd(q, '"Region"', '"Easter"')]),
    gpdCase('column-renamed', 'edit/column-item', [gpd(q, '"Flag"', 'FALSE'), gpd(q, '"Flag"', '"FALSE"'), gpd(q, '"Flag"', '"No"'), gpd(q, '"Flag"', '"no"'), gpd(q, '"Region"', '"East"', '"Flag"', 'FALSE'), gpd(q, '"Region"', '"East"', '"Flag"', '"No"'), gpd(q, '"Flag"', 'TRUE')]),
    gpdCase('row-renamed-number', 'edit/row-item-number', [gpd(q, '"Region"', '"East"'), gpd(q, '"Region"', '"7"'), gpd(q, '"Region"', '7'), gpd(q, '"Region"', '"North"')]),
    renamed('collapsed-tabular', 'gpd-held-2/collapsed-some-tabular', 'A4', 'other', collapsed),
    renamed('collapsed-compact', 'gpd-held-2/collapsed-all-compact', 'A4', 'other', collapsed),
  ]
}
