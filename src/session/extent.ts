/**
 * Where a sheet's cells end (`Session.getUsedRange`): found from what the engine holds, the sheet's columns of values, its
 * spills, its input cells without a value, its looks and its merged areas and tables, never by reading every cell.
 */
import type { SheetExtent } from '../contract.js'
import type { EngineState } from '../engine.js'
import { parseMerge } from '../layout/merges.js'
import { layerOf } from '../looks/sheet.js'
import { MAX_COLS } from '../model.js'

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
  for (const k of state.bare.get(s) ?? []) reach(Math.floor(k / MAX_COLS), k % MAX_COLS)
  // the cells of a look of the sheet's looks, as cells of a look alone
  const looked = layerOf(state.input.sheets[s])?.bounds()
  if (looked) reach(looked.r2, looked.c2)
  for (const m of state.input.sheets[s].merges ?? []) {
    const x = parseMerge(m)
    if (x) reach(x.r2, x.c2)
  }
  for (const t of state.tables.list) if (t.s === s) reach(t.r2, t.c2)
  return { rows, columns }
}
