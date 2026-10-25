/**
 * Where a sheet's cells end (`Session.getUsedRange`): found from what the engine holds, the sheet's columns of values, its
 * spills, its input cells without a value, its looks and its merged areas and tables, never by reading every cell.
 */
import type { SheetExtent, SheetInput } from '../contract.js'
import type { EngineState } from '../engine.js'
import { parseMerge } from '../layout/merges.js'
import { layerOf } from '../looks/sheet.js'
import { colOf, rowOf } from '../model.js'
import type { Tables } from '../tables/model.js'

/** What a sheet's layout reaches: its merged areas, its tables on the sheet, and its looks' bounds (`sheetExtent`, a saved sheet). */
export function reachLayout(sheet: SheetInput, tables: Tables, s: number, reach: (r: number, c: number) => void): void {
  for (const m of sheet.merges ?? []) {
    const x = parseMerge(m)
    if (x) reach(x.r2, x.c2)
  }
  for (const t of tables.list) if (t.s === s) reach(t.r2, t.c2)
  const looked = layerOf(sheet)?.bounds()
  if (looked) reach(looked.r2, looked.c2)
}

export function sheetExtent(state: EngineState, s: number): SheetExtent {
  let rows = 0
  let columns = 0
  const reach = (r: number, c: number) => {
    if (r >= rows) rows = r + 1
    if (c >= columns) columns = c + 1
  }
  // values: input values, formulas and the stored cells of spills; from the right, each column's last value, and once a
  // column with one is found, only the values below the rows reached so far
  const cols = state.ctx.sheets[s].cols
  for (let c = cols.length - 1; c >= 0; c--) {
    const col = cols[c]
    if (!col) continue
    const floor = columns ? rows : 0
    let r = col.length - 1
    while (r >= floor && (col.get(r) ?? null) === null) r--
    if (r >= floor) reach(r, c)
  }
  // the whole range of each spill, its cells past the stored ones (`spillFill`) too
  for (const x of state.spills.on.values()) if (x.s === s) reach(x.r2, x.c2)
  for (const k of state.bare.get(s) ?? []) reach(rowOf(k), colOf(k))
  // the sheet's layout: its looks, merged areas and tables
  reachLayout(state.input.sheets[s], state.tables, s, reach)
  return { rows, columns }
}
