/**
 * A session's edits of the cells of a PivotTable's area, refused as Excel refuses an entry there: a value, a formula or
 * a clear of a cell of the area or of its Filters area (not of the blank row between them). A cell's style, number
 * format, note or link may change.
 */
import { CONTENT } from '../array-areas.js'
import type { CellInput, CellReplace, PivotRefusalReason, PivotTable } from '../contract.js'
import { sheetIndexOf, type EngineState } from '../engine.js'
import { PivotInputError } from '../input-error.js'
import { parseAddress } from '../model.js'
import { hasCell, parseRange } from '../structure/area.js'
import { pivotParts } from './area.js'

/** Excel's alert of a refusal of a PivotTable's cells by its message (`Refusal.reason`); the other refusals have none. */
export const PIVOT_REASONS: Partial<Record<string, PivotRefusalReason>> = { pivotPart: 'part', pivotFormula: 'formula', pivotClear: 'affect' }

/** The input changes the cell's content: its value, formula or array. */
const changes = (old: CellInput | undefined, input: CellInput | null) => CONTENT.some((k) => JSON.stringify(old?.[k]) !== JSON.stringify(input?.[k]))

/** The first edit of a cell of a PivotTable's area that changes its content: the sheet, the cell, the PivotTable and the input; null when none does. */
export function pivotPart(state: EngineState, edits: CellReplace[]): { sheet: string; cell: string; pivot: string; table: PivotTable; input: CellInput | null } | null {
  const sheets = state.input.sheets
  if (!sheets.some((s) => s.pivotTables?.length)) return null
  for (const e of edits) {
    const sh = sheets[sheetIndexOf(state, e.sheet) ?? -1]
    if (!sh?.pivotTables?.length || !changes(sh.cells[e.cell], e.input)) continue
    const { row, col } = parseAddress(e.cell)
    const t = sh.pivotTables.find((p) => pivotParts(p).some((x) => hasCell(x, row, col)))
    if (t) return { sheet: sh.name, cell: e.cell, pivot: t.name, table: t, input: e.input }
  }
  return null
}

/**
 * Edits of a session that change a cell of a PivotTable's area throw `PivotInputError` (`pivotPart`) with Excel's alert:
 * a cell cleared affects the PivotTable, a formula over a label is one for an item or field name, the rest change its part.
 */
export function checkPivotEdits(state: EngineState, edits: CellReplace[]) {
  const p = pivotPart(state, edits)
  if (!p) return
  const { input } = p
  const { row, col } = parseAddress(p.cell)
  const data = p.table.location.areas?.data
  const value = data !== undefined && hasCell(parseRange(data)!, row, col)
  const reason = input?.value === undefined && input?.formula === undefined ? 'affect' : input.formula !== undefined && !value ? 'formula' : 'part'
  throw new PivotInputError('edits', `${p.sheet}!${p.cell} is part of the PivotTable ${p.pivot}: Excel does not let an entry change a PivotTable's cells`, reason, p.sheet, p.cell)
}
