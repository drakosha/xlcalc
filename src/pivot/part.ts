/**
 * A session's edits of the cells of a PivotTable's area, refused as Excel refuses an entry there: a value, a formula or
 * a clear of a cell of the area or of its Filters area (not of the blank row between them). A cell's style, number
 * format, note or link may change.
 */
import { CONTENT } from '../array-areas.js'
import type { CellInput, CellReplace } from '../contract.js'
import { sheetIndexOf, type EngineState } from '../engine.js'
import { InputError } from '../input-error.js'
import { parseAddress } from '../model.js'
import { hasCell } from '../structure/area.js'
import { pivotParts } from './area.js'

/** The input changes the cell's content: its value, formula or array. */
const changes = (old: CellInput | undefined, input: CellInput | null) => CONTENT.some((k) => JSON.stringify(old?.[k]) !== JSON.stringify(input?.[k]))

/** The first edit of a cell of a PivotTable's area that changes its content: the sheet, the cell and the PivotTable; null when none does. */
export function pivotPart(state: EngineState, edits: CellReplace[]): { sheet: string; cell: string; pivot: string } | null {
  const sheets = state.input.sheets
  if (!sheets.some((s) => s.pivotTables?.length)) return null
  for (const e of edits) {
    const sh = sheets[sheetIndexOf(state, e.sheet) ?? -1]
    if (!sh?.pivotTables?.length || !changes(sh.cells[e.cell], e.input)) continue
    const { row, col } = parseAddress(e.cell)
    const t = sh.pivotTables.find((p) => pivotParts(p.location).some((x) => hasCell(x, row, col)))
    if (t) return { sheet: sh.name, cell: e.cell, pivot: t.name }
  }
  return null
}

/** Edits of a session that change a cell of a PivotTable's area throw `InputError` (`pivotPart`). */
export function checkPivotEdits(state: EngineState, edits: CellReplace[]) {
  const p = pivotPart(state, edits)
  if (p) throw new InputError('edits', `${p.sheet}!${p.cell} is part of the PivotTable ${p.pivot}: Excel does not let an entry change a PivotTable's cells`)
}
