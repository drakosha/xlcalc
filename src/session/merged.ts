/**
 * A merge or an unmerge in a session without building the book again. The operation runs as `editStructure` runs it
 * (`applyOp`), on the book with the sheet's cells cut down to those of the merged range, which is all a merge reads
 * when it meets no array formula of the old kind; when it moves no cell, what it gave differs from the session's book in
 * the cells of the range, the sheet's merged areas and the style table only, which are written as a patch of looks is
 * (`applyPlan`).
 */
import type { BookInput, CellInput, CellReplace, Refusal, SheetInput, StructureOp, StructureRefusal } from '../contract.js'
import type { EngineState } from '../engine.js'
import { mergesOf } from '../layout/merges.js'
import { inputIn } from '../looks/sheet.js'
import { formatAddress } from '../model.js'
import { cellCount, meets, parseRange } from '../structure/area.js'
import { applyOp } from '../structure/index.js'
import { grown } from '../structure/merge.js'
import { stateValues } from '../structure/values.js'
import { sheetKey } from '../syntax.js'
import type { Plan } from './styles.js'

/** Past this many cells a merged range is built again with the book, as other operations are. */
const MAX_CELLS = 65_536

const same = (a: CellInput | undefined, b: CellInput | undefined) => a === b || JSON.stringify(a ?? null) === JSON.stringify(b ?? null)

/**
 * The plan of a merge or an unmerge, its refusal, or null when the book is to be built again: a range too large or
 * meeting an array formula of the old kind, a cell moved into an empty top-left cell, input the full operation checks.
 */
export function mergedPlan(s: EngineState, op: Extract<StructureOp, { op: 'mergeCells' | 'unmergeCells' }>): Plan | Refusal<StructureRefusal> | null {
  const i = typeof op.sheet === 'string' ? s.input.sheets.findIndex((x) => sheetKey(x.name) === sheetKey(op.sheet)) : -1
  const range = i < 0 || typeof op.range !== 'string' ? null : parseRange(op.range)
  if (!range) return null
  const was = s.input.sheets[i]
  const area = op.op === 'mergeCells' ? grown(range, mergesOf(was)) : null
  if (area && (cellCount(area) > MAX_CELLS || s.formulas.some((f) => f?.area && f.s === i && meets(f.area, area)))) return null
  const cells: Record<string, CellInput> = {}
  if (area)
    for (let r = area.r1; r <= area.r2; r++)
      for (let c = area.c1; c <= area.c2; c++) {
        const a = formatAddress(r, c)
        if (a in was.cells) cells[a] = was.cells[a]
      }
  const light: BookInput = { ...s.input, sheets: s.input.sheets.map((x, k) => (k === i ? { ...was, cells } : x)) }
  const r = applyOp(light, op, stateValues(s))
  if (!r.ok) return r
  if (r.geometry.length) return null
  const now = r.book.sheets[i]
  const edits: CellReplace[] = []
  if (area)
    for (let r = area.r1; r <= area.r2; r++)
      for (let c = area.c1; c <= area.c2; c++) {
        const a = formatAddress(r, c)
        const input = inputIn(now, a, r, c)
        if (!same(inputIn(was, a, r, c), input)) edits.push({ sheet: was.name, cell: a, input: input ?? null })
      }
  const sheet: SheetInput = { ...was, merges: now.merges ?? [] }
  if (!now.merges) delete sheet.merges
  return {
    edits,
    sheets: new Map([[i, sheet]]),
    touched: [{ sheet: was.name, merges: now.merges ?? [] }],
    ...(r.book.styles !== s.input.styles ? { styles: r.book.styles } : {}),
  }
}
