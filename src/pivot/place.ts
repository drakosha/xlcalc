/**
 * PivotTables laid out written to their sheets as one step (a refresh of a cache shared by several): every new area
 * checked first, then the old areas cleared and the new cells written, the definitions with their new layout and place.
 * An area over cells that hold something, a table or another PivotTable refuses the whole step, and nothing changes.
 */
import type { BookInput, CellInput, PivotTable, SheetInput } from '../contract.js'
import { CONTENT, holds } from '../array-areas.js'
import { say } from '../locale/messages.js'
import { formatAddress, readAddress } from '../model.js'
import { cornersText, hasCell, meets, parseRange, type Rect } from '../structure/area.js'
import type { Applied } from '../structure/geometry.js'
import { sheetIndex } from '../structure/rewrite.js'
import { ERROR_TEXT } from '../syntax/errors.js'
import { pivotArea } from './area.js'
import type { PivotLaidOut } from './layout.js'

/** A PivotTable to write: its sheet, its definition as refreshed (its `name` the one on the sheet), and its layout. */
export interface Placement {
  sheet: string
  table: PivotTable
  laid: PivotLaidOut
}

/** A cell's fields that a PivotTable's layout writes: its content and number format; its style, note and link stay. */
const LAID = [...CONTENT, 'numberFormat'] as const

const cleared = (c: CellInput): CellInput | undefined => {
  const out: CellInput = { ...c }
  for (const k of LAID) delete out[k]
  return Object.keys(out).length ? out : undefined
}

/** Writes the PivotTables `placed` to their sheets at once (`Placement`), or refuses (`pivot`) when an area cannot be laid. */
export function placePivots(book: BookInput, placed: Placement[]): Applied {
  const steps = placed.map((p) => {
    const s = sheetIndex(book, p.sheet)
    const pivots = book.sheets[s]?.pivotTables ?? []
    const at = pivots.findIndex((t) => t.name === p.table.name)
    if (at < 0) throw new Error(`no PivotTable ${p.table.name} on ${p.sheet}`)
    return { ...p, s, at, old: pivotArea(pivots[at].location) }
  })
  // each new area against what lies there that is not a PivotTable laid again: cells, tables, other PivotTables
  for (const p of steps) {
    const sheet = book.sheets[p.s]
    const x = p.laid.area
    const fail = (key: 'pivotOverPivot' | 'pivotOverTable' | 'pivotOverData', more: Record<string, string>): Applied => ({ ok: false, code: 'pivot', message: say(book.locale, { key, pivot: p.table.name, ...more } as never) })
    for (const q of steps) if (q !== p && q.s === p.s && meets(x, q.laid.area)) return fail('pivotOverPivot', { other: q.table.name })
    for (const [i, t] of (sheet.pivotTables ?? []).entries()) {
      const area = steps.some((q) => q.s === p.s && q.at === i) ? null : pivotArea(t.location)
      if (area && meets(x, area)) return fail('pivotOverPivot', { other: t.name })
    }
    for (const t of sheet.tables ?? []) {
      const r = parseRange(t.range)
      if (r && meets(x, r)) return fail('pivotOverTable', { table: t.name })
    }
    const own = steps.flatMap((q) => (q.s === p.s && q.old ? [q.old] : []))
    const hit = filled(sheet, x, own)
    if (hit) return fail('pivotOverData', { sheet: sheet.name, range: cornersText(x) })
  }
  const sheets = [...book.sheets]
  for (const p of steps) {
    const sheet = sheets[p.s]
    const cells: SheetInput['cells'] = { ...sheet.cells }
    if (p.old) for (const [addr, c] of cellsIn(cells, p.old)) {
      const rest = cleared(c)
      if (rest) cells[addr] = rest
      else delete cells[addr]
    }
    for (const c of p.laid.cells) {
      const addr = formatAddress(c.row, c.col)
      const out: CellInput = { ...cells[addr] }
      if (c.error) out.formula = ERROR_TEXT[c.error]
      else if (c.value !== undefined) out.value = c.value
      if (c.numberFormat !== undefined) out.numberFormat = c.numberFormat
      cells[addr] = out
    }
    const pivotTables = [...(sheet.pivotTables ?? [])]
    pivotTables[p.at] = { ...p.table, location: p.laid.location, rowLines: p.laid.rowLines, columnLines: p.laid.columnLines }
    sheets[p.s] = { ...sheet, cells, pivotTables }
  }
  return { ok: true, book: { ...book, sheets }, geometry: [] }
}

/** The cells of `x` among `cells`, with their addresses. */
function cellsIn(cells: SheetInput['cells'], x: Rect): [string, CellInput][] {
  const at = { r: 0, c: 0 }
  return Object.entries(cells).filter(([addr]) => {
    readAddress(addr, at)
    return hasCell(x, at.r, at.c)
  })
}

/** Whether a cell of `x` outside the areas `own` holds a value or a formula. */
function filled(sheet: SheetInput, x: Rect, own: Rect[]): boolean {
  const at = { r: 0, c: 0 }
  for (const [addr, c] of cellsIn(sheet.cells, x)) {
    readAddress(addr, at)
    if (holds(c) && !own.some((o) => hasCell(o, at.r, at.c))) return true
  }
  return false
}
