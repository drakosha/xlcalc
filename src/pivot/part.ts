/**
 * A session's edits of the cells of a PivotTable's area, refused as Excel refuses an entry there: a value, a formula or
 * a clear of a cell of the area or of its Filters area (not of the blank row between them). A cell's style, number
 * format, note or link may change.
 */
import { CONTENT } from '../array-areas.js'
import type { CellInput, CellReplace, PivotRefusalReason, PivotTable, RangeRef } from '../contract.js'
import { sheetIndexOf, type EngineState } from '../engine.js'
import { PivotInputError } from '../input-error.js'
import { parseAddress } from '../model.js'
import { hasCell, parseRange, type Rect } from '../structure/area.js'
import { pivotParts } from './area.js'

/** Excel's alert of a refusal of a PivotTable's cells by its message (`Refusal.reason`); the other refusals have none. */
export const PIVOT_REASONS: Partial<Record<string, PivotRefusalReason>> = { pivotPart: 'part', pivotFormula: 'formula', pivotAffect: 'affect' }

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

/** The rows and columns of `x` all covered by `rects`: each row's covering columns merged run over its width. */
function covered(x: Rect, rects: Rect[]): boolean {
  for (let r = x.r1; r <= x.r2; r++) {
    const spans = rects.filter((y) => y.r1 <= r && y.r2 >= r && y.c2 >= x.c1 && y.c1 <= x.c2).sort((a, b) => a.c1 - b.c1)
    let next = x.c1
    for (const y of spans) if (y.c1 <= next) next = Math.max(next, y.c2 + 1)
    if (next <= x.c2) return false
  }
  return true
}

/**
 * The PivotTables whose every cell the ranges cover (their table and Filters area, not the blank row between): Excel
 * deletes such a PivotTable when its contents are cleared (the set alerts), where a clear of part of it is refused.
 */
export function clearedPivots(state: EngineState, ranges: RangeRef[]): { sheet: string; table: PivotTable }[] {
  const sheets = state.input.sheets
  if (!sheets.some((s) => s.pivotTables?.length)) return []
  return sheets.flatMap((sh, i) => {
    const rects = ranges.filter((x) => sheetIndexOf(state, x.sheet) === i).map((x) => parseRange(x.range)!)
    return rects.length ? (sh.pivotTables ?? []).filter((t) => pivotParts(t).every((x) => covered(x, rects))).map((table) => ({ sheet: sh.name, table })) : []
  })
}

/** The edits of cells outside the PivotTables `gone`, which a clear deletes. */
export function outsidePivots(state: EngineState, edits: CellReplace[], gone: { sheet: string; table: PivotTable }[]): CellReplace[] {
  return edits.filter((e) => {
    const { row, col } = parseAddress(e.cell)
    const i = sheetIndexOf(state, e.sheet)
    return !gone.some((p) => sheetIndexOf(state, p.sheet) === i && pivotParts(p.table).some((x) => hasCell(x, row, col)))
  })
}

/** The cell at row `r`, column `c` of sheet `i` is a PivotTable's part: of its table or its Filters area. */
export const inPivot = (state: EngineState, i: number, r: number, c: number): boolean => (state.input.sheets[i]?.pivotTables ?? []).some((t) => pivotParts(t).some((x) => hasCell(x, r, c)))
