/**
 * Where a sheet's cells end (`Session.getUsedRange`): found from what the engine holds, the sheet's columns of values, its
 * spills, its input cells without a value, its looks and its merged areas and tables, never by reading every cell.
 */
import type { SheetExtent, SheetInput } from '../contract.js'
import type { EngineState } from '../engine.js'
import { parseMerge } from '../layout/merges.js'
import { layerOf } from '../looks/sheet.js'
import { colOf, readAddress, rowOf } from '../model.js'
import type { Rect } from '../structure/area.js'
import type { Tables } from '../tables/model.js'

/** The rectangles a sheet's layout covers: its merged areas, its tables on the sheet, and its looks' bounds (`sheetExtent`, a saved sheet, `usedArea`). */
export function reachLayout(sheet: SheetInput, tables: Tables, s: number, reach: (x: Rect) => void): void {
  for (const m of sheet.merges ?? []) {
    const x = parseMerge(m)
    if (x) reach(x)
  }
  for (const t of tables.list) if (t.s === s) reach(t)
  const looked = layerOf(sheet)?.bounds()
  if (looked) reach(looked)
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
  reachLayout(state.input.sheets[s], state.tables, s, (x) => reach(x.r2, x.c2))
  return { rows, columns }
}

/**
 * Excel's used range of a sheet: from the first row and column to the last ones of its cells (values, formulas, input
 * cells without a value), its spills and its layout (`sheetExtent`); null for an empty sheet.
 */
export function usedArea(state: EngineState, s: number): Rect | null {
  const { rows, columns } = sheetExtent(state, s)
  if (!rows) return null
  const x: Rect = { r1: rows - 1, c1: columns - 1, r2: rows - 1, c2: columns - 1 }
  const reach = (r: number, c: number) => {
    if (r < x.r1) x.r1 = r
    if (c < x.c1) x.c1 = c
  }
  const at = { r: 0, c: 0 }
  for (const a of Object.keys(state.input.sheets[s].cells)) {
    readAddress(a, at)
    reach(at.r, at.c)
  }
  for (const y of state.spills.on.values()) if (y.s === s) reach(y.r1, y.c1)
  reachLayout(state.input.sheets[s], state.tables, s, (y) => reach(y.r1, y.c1))
  return x
}
