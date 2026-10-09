/**
 * A copy from a filtered sheet over rows a filter hid (the struct set `hidden`): Excel copies the visible rows only,
 * packed one after another, and pastes them as values, as it pastes a selection of several areas. The copy is made
 * from a scratch sheet holding them so; a selection of hidden rows only is copied whole.
 */
import type { BookInput, CellInput, RowInput, SheetInput } from '../contract.js'
import { spend } from '../exec/budget.js'
import { entryAt, lines, toMap, type Line } from '../layout/spans.js'
import { formatAddress, parseAddress } from '../model.js'
import { ERROR_TEXT, sheetKey } from '../syntax.js'
import type { Rect } from './area.js'
import { visibleRows } from './filtered.js'
import type { ValuesIn } from './values.js'

/** The scratch sheet's name, which no sheet of a book can have. */
export const CLIP = '\u0000clip'

/** A computed value as a cell to paste: an error as a formula of it, as an xlsx file's error constant is read. */
const pasted = (v: NonNullable<ReturnType<ValuesIn>[number][2]>): CellInput => (typeof v === 'object' && v !== null ? { formula: ERROR_TEXT[v.error] } : { value: v })

/** The book with the scratch sheet and the copy's source on it, or null when the copy leaves out no row. */
export function clipped(book: BookInput, from: string, src: Rect, values: ValuesIn): { book: BookInput; src: Rect } | null {
  const visible = visibleRows(book, from, src.r1, src.r2)
  if (!visible?.length) return null
  const sheet = book.sheets.find((s) => sheetKey(s.name) === sheetKey(from))!
  const row = new Map(visible.map((r, k) => [r, src.r1 + k]))
  const cells: Record<string, CellInput> = {}
  // looks and meta of the cells, then the values over them
  for (const addr of Object.keys(sheet.cells)) {
    const cell = sheet.cells[addr]
    spend(1)
    const p = parseAddress(addr)
    const r = row.get(p.row)
    if (r === undefined || p.col < src.c1 || p.col > src.c2) continue
    const { value: _v, formula: _f, ...look } = cell
    if (Object.keys(look).length) cells[formatAddress(r, p.col)] = look
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
