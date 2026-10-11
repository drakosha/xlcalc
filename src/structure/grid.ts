/**
 * A sheet's cells read once for a structure operation: each address parsed once, in the sheet's order. The steps of one
 * operation (`withGrids`) share it: its refusals, the cells that move and those they land on, the line inserted lines
 * copy their looks from, a copy's source, the array formulas. A book's sheets do not change while an operation reads
 * them (it makes new ones), so a sheet's cells object stands for its grid until the operation ends.
 */
import type { CellInput, SheetInput } from '../contract.js'
import { spend } from '../exec/budget.js'
import { readAddress } from '../model.js'
import { hasCell, parseRange, type Rect } from './area.js'

export interface Grid {
  addr: string[]
  row: number[]
  col: number[]
  cell: CellInput[]
}

let shared: WeakMap<object, Grid> | null = null
const ARRAYS = new WeakMap<Grid, [string, Rect][]>()

/** Runs one operation with the grids of the sheets it reads shared by its steps. */
export function withGrids<T>(f: () => T): T {
  if (shared) return f()
  shared = new WeakMap()
  try {
    return f()
  } finally {
    shared = null
  }
}

/** A sheet's cells with their rows and columns, read in one pass (once for an operation under `withGrids`). */
export function gridOf(sheet: SheetInput): Grid {
  const known = shared?.get(sheet.cells)
  if (known) return known
  const g: Grid = { addr: [], row: [], col: [], cell: [] }
  const p = { r: 0, c: 0 }
  for (const addr in sheet.cells) {
    // a session's operation stops here at its deadline or abort (`exec/budget.ts`)
    spend(1)
    readAddress(addr, p)
    g.addr.push(addr)
    g.row.push(p.r)
    g.col.push(p.c)
    g.cell.push(sheet.cells[addr])
  }
  shared?.set(sheet.cells, g)
  return g
}

/** The indices of a sheet's cells that lie in a rectangle, in the sheet's order. */
export function cellsIn(g: Grid, x: Rect): number[] {
  const out: number[] = []
  for (let i = 0; i < g.addr.length; i++) if (hasCell(x, g.row[i], g.col[i])) out.push(i)
  return out
}

/** The array formulas of the old kind of a sheet (`CellInput.array`): their cell and area. */
export function arraysOf(sheet: SheetInput): [string, Rect][] {
  const g = gridOf(sheet)
  let out = ARRAYS.get(g)
  if (!out) {
    out = []
    for (let i = 0; i < g.cell.length; i++) if (g.cell[i].array !== undefined && g.cell[i].formula !== undefined) out.push([g.addr[i], parseRange(g.cell[i].array!)!])
    ARRAYS.set(g, out)
  }
  return out
}
