/**
 * A copy from a filtered sheet over rows a filter hid (the struct set `hidden`): Excel copies the visible rows only,
 * packed one after another, and pastes them as values, as it pastes a selection of several areas. The copy is made
 * from a scratch sheet holding them so; a selection of hidden rows only is copied whole.
 */
import type { BookInput, CellInput, RowInput, SheetInput } from '../contract.js'
import { entryAt, lines, toMap, type Line } from '../layout/spans.js'
import { formatAddress } from '../model.js'
import { ERROR_TEXT } from '../syntax.js'
import type { Rect } from './area.js'
import { visibleRows } from './filtered.js'
import { cellsIn, gridOf } from './grid.js'
import { sheetOf } from './rewrite.js'
import type { ValuesIn } from './values.js'

/** The scratch sheet's name, which no sheet of a book can have. */
export const CLIP = '\u0000clip'

/** A computed value as a cell to paste: an error as a formula of it, as an xlsx file's error constant is read. */
const pasted = (v: NonNullable<ReturnType<ValuesIn>[number][2]>): CellInput => (typeof v === 'object' && v !== null ? { formula: ERROR_TEXT[v.error] } : { value: v })

/** The book with the scratch sheet and the copy's source on it, or null when the copy leaves out no row. */
export function clipped(book: BookInput, from: string, src: Rect, values: ValuesIn): { book: BookInput; src: Rect } | null {
  const visible = visibleRows(book, from, src.r1, src.r2)
  if (!visible?.length) return null
  const sheet = sheetOf(book, from)
  const row = new Map(visible.map((r, k) => [r, src.r1 + k]))
  const cells: Record<string, CellInput> = {}
  // looks and meta of the cells, then the values over them
  const g = gridOf(sheet)
  for (const i of cellsIn(g, src)) {
    const r = row.get(g.row[i])
    if (r === undefined) continue
    const { value: _v, formula: _f, ...look } = g.cell[i]
    if (Object.keys(look).length) cells[formatAddress(r, g.col[i])] = look
  }
  for (const [r, c, v] of values(sheet.name, src)) {
    const at = row.get(r)
    if (at === undefined || v === null) continue
    const addr = formatAddress(at, c)
    cells[addr] = { ...cells[addr], ...pasted(v) }
  }
  const rows: Line<RowInput>[] = []
  const sheetRows = lines(sheet.rows, 'r')
  for (const [r, at] of row) {
    const e = entryAt(sheetRows, r)
    if (e) rows.push([at, at, e])
  }
  const clip: SheetInput = { name: CLIP, cells, ...(rows.length ? { rows: toMap(rows, 'r') } : {}), ...(sheet.columns ? { columns: sheet.columns } : {}) }
  return { book: { ...book, sheets: [...book.sheets, clip] }, src: { ...src, r2: src.r1 + visible.length - 1 } }
}
