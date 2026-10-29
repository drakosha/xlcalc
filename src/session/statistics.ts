/**
 * Excel's Workbook Statistics (`Session.getWorkbookStatistics`, the set `bench/coverage/stats`): a sheet's end, the
 * last cell of its used range, and the counts of a sheet and of the workbook, found from the book's input and its
 * spills without reading every cell. A cell with data holds a value or a formula; each cell a spilled array or a legacy
 * array formula fills is one with data and one with a formula, as Excel counts them. What the workbook does not hold is null.
 */
import type { StatisticsCounts, WorkbookStatistics } from '../contract.js'
import type { EngineState } from '../engine.js'
import { arrayArea, holds } from '../array-areas.js'
import { formatAddress } from '../model.js'
import type { Rect } from '../structure/area.js'
import { usedArea } from './extent.js'

/** The counts of what the workbook does not hold. */
const UNHELD = { pivotTables: null, charts: null, images: null, formControls: null, objects: null, comments: null } as const

const area = (x: Rect) => (x.r2 - x.r1 + 1) * (x.c2 - x.c1 + 1)

/** A sheet's counts: its input cells, the cells its arrays fill past their formulas, its tables and notes. */
function counts(state: EngineState, s: number): StatisticsCounts {
  let [cells, formulas, notes] = [0, 0, 0]
  // the cells an array fills past its formula's
  const fill = (x: Rect) => {
    cells += area(x) - 1
    formulas += area(x) - 1
  }
  for (const [addr, c] of Object.entries(state.input.sheets[s].cells)) {
    if (holds(c)) cells++
    if (c.note) notes++
    if (c.formula === undefined) continue
    formulas++
    const x = c.array === undefined ? null : arrayArea(addr, c.array)
    if (x) fill(x)
  }
  for (const x of state.spills.on.values()) if (x.s === s) fill(x)
  const tables = state.tables.list.filter((t) => t.s === s).length
  return { cells, tables, formulas, notes, ...UNHELD }
}

export function workbookStatistics(state: EngineState, s: number): WorkbookStatistics {
  const used = usedArea(state, s)
  const book = state.input.sheets.map((_, i) => counts(state, i))
  const sum = (k: 'cells' | 'tables' | 'formulas' | 'notes') => book.reduce((n, x) => n + x[k], 0)
  return {
    sheet: { end: used ? formatAddress(used.r2, used.c2) : 'A1', ...book[s] },
    workbook: { sheets: book.length, cells: sum('cells'), tables: sum('tables'), formulas: sum('formulas'), notes: sum('notes'), ...UNHELD },
  }
}
