/**
 * How a structure operation moved a book's addresses (`EditResult.moved`), done or taken back: lines inserted and
 * deleted, a range cut and pasted, a sheet renamed, deleted, inserted or moved. A host moves its window by them.
 */
import type { AddressMove, StructureOp } from '../contract.js'
import { colToIndex } from '../model.js'
import { parseRange } from '../structure/area.js'
import { sheetKey } from '../syntax.js'
import { areaText } from '../validation/rules.js'
import type { Way } from './rebook.js'

/**
 * The moves of `op` done (`forth`) or taken back (`back`), from the book whose sheets are `before` to the one whose sheets
 * are `after`; none for an operation that moves no address (a copy, a table, a sheet hidden or shown).
 */
export function movesOf(op: StructureOp | null, way: Way, before: string[], after: string[]): AddressMove[] {
  if (!op) return []
  const back = way === 'back'
  const spelled = (name: string, sheets = after) => sheets.find((n) => sheetKey(n) === sheetKey(name)) ?? name
  const index = (name: string) => after.findIndex((n) => sheetKey(n) === sheetKey(name))
  switch (op.op) {
    case 'insertRows':
    case 'deleteRows':
    case 'insertColumns':
    case 'deleteColumns': {
      const rows = op.op === 'insertRows' || op.op === 'deleteRows'
      const at = rows ? (op as { row: number }).row : colToIndex((op as { column: string }).column) + 1
      const count = (op.count ?? 1) * (op.op.startsWith('insert') === back ? -1 : 1)
      return [{ kind: 'lines', sheet: spelled(op.sheet), axis: rows ? 'rows' : 'columns', at, count }]
    }
    case 'moveRange': {
      const src = parseRange(op.from.range)!
      const to = parseRange(op.to.cell)!
      const [dr, dc] = [to.r1 - src.r1, to.c1 - src.c1]
      if (!back) return [{ kind: 'range', from: { sheet: spelled(op.from.sheet), range: areaText(src) }, to: { sheet: spelled(op.to.sheet), cell: op.to.cell } }]
      const landed = { r1: src.r1 + dr, r2: src.r2 + dr, c1: src.c1 + dc, c2: src.c2 + dc }
      return [{ kind: 'range', from: { sheet: spelled(op.to.sheet, before), range: areaText(landed) }, to: { sheet: spelled(op.from.sheet), cell: areaText({ ...src, r2: src.r1, c2: src.c1 }) } }]
    }
    case 'renameSheet':
      return back ? [{ kind: 'renameSheet', sheet: spelled(op.name, before), name: spelled(op.sheet) }] : [{ kind: 'renameSheet', sheet: spelled(op.sheet, before), name: op.name }]
    case 'deleteSheet':
      return back ? [{ kind: 'insertSheet', sheet: spelled(op.sheet), index: index(op.sheet) }] : [{ kind: 'deleteSheet', sheet: spelled(op.sheet, before) }]
    case 'insertSheet':
      return back ? [{ kind: 'deleteSheet', sheet: spelled(op.name, before) }] : [{ kind: 'insertSheet', sheet: spelled(op.name), index: index(op.name) }]
    case 'moveSheet':
      return [{ kind: 'moveSheet', sheet: spelled(op.sheet), index: index(op.sheet) }]
    default:
      return []
  }
}
