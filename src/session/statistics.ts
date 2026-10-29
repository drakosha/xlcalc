/**
 * Excel's Workbook Statistics (`Session.getWorkbookStatistics`, the sets `bench/coverage/stats`): the counts of a sheet
 * and of the workbook and a sheet's end, found from the book's input and its spills without reading every cell. A cell
 * with data holds a value or a formula; each cell a spilled array or a legacy array formula fills is one with data and
 * one with a formula, as Excel counts them. The end is the last row and the last column of the cells with data or a
 * note, not of the used range: a style, a row or column of its own, a hyperlink alone or a table does not reach it.
 * What the workbook does not hold is null.
 */
import type { StatisticsCounts, WorkbookStatistics } from '../contract.js'
import type { EngineState } from '../engine.js'
import { arrayArea, holds } from '../array-areas.js'
import { formatAddress, readAddress } from '../model.js'
import type { Rect } from '../structure/area.js'

/** The counts of what the workbook does not hold. */
const UNHELD = { pivotTables: null, charts: null, images: null, formControls: null, objects: null, comments: null } as const

const area = (x: Rect) => (x.r2 - x.r1 + 1) * (x.c2 - x.c1 + 1)

/** A sheet's counts: its input cells, the cells its arrays fill past their formulas, its tables and notes; and its end. */
function counts(state: EngineState, s: number): StatisticsCounts & { end: string | null } {
  let [cells, formulas, notes, rows, cols] = [0, 0, 0, 0, 0]
  const reach = (r: number, c: number) => {
    if (r >= rows) rows = r + 1
    if (c >= cols) cols = c + 1
  }
  // the cells an array fills past its formula's
  const fill = (x: Rect) => {
    cells += area(x) - 1
    formulas += area(x) - 1
    reach(x.r2, x.c2)
  }
  const at = { r: 0, c: 0 }
  for (const [addr, c] of Object.entries(state.input.sheets[s].cells)) {
    const data = holds(c)
    if (data) cells++
    if (c.note) notes++
    if (data || c.note) {
      readAddress(addr, at)
      reach(at.r, at.c)
    }
    if (c.formula === undefined) continue
    formulas++
    const x = c.array === undefined ? null : arrayArea(addr, c.array)
    if (x) fill(x)
  }
  for (const x of state.spills.on.values()) if (x.s === s) fill(x)
  const tables = state.tables.list.filter((t) => t.s === s).length
  return { end: rows ? formatAddress(rows - 1, cols - 1) : null, cells, tables, formulas, notes, ...UNHELD }
}

export function workbookStatistics(state: EngineState, s: number): WorkbookStatistics {
  const book = state.input.sheets.map((_, i) => counts(state, i))
  const sum = (k: 'cells' | 'tables' | 'formulas' | 'notes') => book.reduce((n, x) => n + x[k], 0)
  return {
    sheet: book[s],
    workbook: { sheets: book.length, cells: sum('cells'), tables: sum('tables'), formulas: sum('formulas'), notes: sum('notes'), ...UNHELD },
  }
}
