/**
 * Excel's Merge Cells and Unmerge Cells (`mergeCells`, `unmergeCells`). A merge takes in whole each merged area its range
 * meets, growing until it meets no other; the range then becomes one merged area, or one a row with `across`. Each new
 * area keeps its top-left cell; when that is empty, the first cell holding a value or a formula, row by row, is moved
 * there as cut and paste moves it (formulas reading it follow, its look comes along). The other cells lose their values
 * and formulas and take the top-left cell's style and number format (its own, else its row's, else its column's), so
 * that they keep it when the area is split again. With
 * `center` the top-left cells' style is centred across. Cells of a table, and of an array formula of the old kind, are
 * not merged. An unmerge drops the merged areas its range meets; the cells keep what they hold and look like.
 */
import type { BookInput, CellInput, LookRange, SheetInput } from '../contract.js'
import { mergesOf, parseMerge } from '../layout/merges.js'
import { lineAt } from '../layout/spans.js'
import { say } from '../locale/messages.js'
import { inputIn } from '../looks/sheet.js'
import { formatAddress } from '../model.js'
import { internStyle } from '../style-table.js'
import { tablesOf } from '../tables/model.js'
import { cornersText, grown, meets, rectText, type Rect } from './area.js'
import type { Applied, Geometry } from './geometry.js'
import { moveRange } from './move.js'
import { arraysOf, cellsIn, gridOf, type Grid } from './grid.js'
import { holds } from '../array-areas.js'

export { grown }

/** The fields of a cell a merge clears: what it holds. */
const CONTENT = ['value', 'formula', 'array', 'held'] as const

export function mergeCells(book: BookInput, sheet: string, x: Rect, across: boolean, center: boolean): Applied {
  const s = book.sheets.findIndex((t) => t.name === sheet)
  const merges = mergesOf(book.sheets[s])
  const area = grown(x, merges)
  const table = tablesOf(book).list.find((t) => t.s === s && meets(t, area))
  if (table) return { ok: false, code: 'table', message: say(book.locale, { key: 'mergeInTable', range: rectText(area), table: table.name }) }
  const array = arraysOf(book.sheets[s]).find(([, a]) => meets(a, area))
  if (array) return { ok: false, code: 'array', message: say(book.locale, { key: 'arrayPart', sheet, cell: array[0], area: rectText(array[1]) }) }
  const areas = across ? Array.from({ length: area.r2 - area.r1 + 1 }, (_, i) => ({ ...area, r1: area.r1 + i, r2: area.r1 + i })) : [area]
  const made = areas.filter((a) => a.r1 < a.r2 || a.c1 < a.c2)
  // the merged areas met are split first, so that a cell of one can move
  const kept = merges.filter((m) => !meets(m, area)).map(cornersText)
  let moved: BookInput = { ...book, sheets: book.sheets.map((t, i) => (i === s ? { ...t, merges: kept } : t)) }
  const geometry: Geometry = []
  // the cells of the area by row, read once for a grid: across, each row's area takes its own (a walk of the grid per row was O(rows × sheet))
  let byRow: { g: Grid; rows: Map<number, number[]> } | undefined
  const cellsOf = (g: Grid, a: Rect): readonly number[] => {
    if (!across) return cellsIn(g, a)
    if (byRow?.g !== g) {
      byRow = { g, rows: new Map() }
      for (const i of cellsIn(g, area)) {
        const list = byRow.rows.get(g.row[i])
        if (list) list.push(i)
        else byRow.rows.set(g.row[i], [i])
      }
    }
    return byRow.rows.get(a.r1) ?? []
  }
  for (const a of made) {
    const g = gridOf(moved.sheets[s])
    if (holds(moved.sheets[s].cells[formatAddress(a.r1, a.c1)])) continue
    const from = cellsOf(g, a)
      .filter((i) => holds(g.cell[i]))
      .sort((i, j) => g.row[i] - g.row[j] || g.col[i] - g.col[j])[0]
    if (from === undefined) continue
    const r = moveRange(moved, sheet, { r1: g.row[from], c1: g.col[from], r2: g.row[from], c2: g.col[from] }, sheet, a.r1, a.c1)
    if (!r.ok) return r
    moved = r.book
    geometry.push(...r.geometry)
  }
  const sh = moved.sheets[s]
  const g = gridOf(sh)
  const cells: Record<string, CellInput> = { ...sh.cells }
  const looks: LookRange[] = []
  let styles = moved.styles ?? []
  for (const a of center ? areas : made) {
    const top = formatAddress(a.r1, a.c1)
    const own = inputIn(sh, top, a.r1, a.c1)
    const [row, col] = [lineAt(sh.rows, 'r', a.r1), lineAt(sh.columns, 'c', a.c1)]
    const numberFormat = own?.numberFormat ?? row?.numberFormat ?? col?.numberFormat
    let style = own?.style ?? row?.style ?? col?.style
    if (center) {
      const was = style === undefined ? {} : styles[style]
      ;({ index: style, styles } = internStyle(styles, { ...was, alignment: { ...was.alignment, horizontal: 'center' } }))
    }
    const look = { ...(style === undefined ? {} : { style }), ...(numberFormat === undefined ? {} : { numberFormat }) }
    // the top-left cell centred holds its look as its own, which the look of the area beneath an entry does not give it
    if (center) cells[top] = { ...withoutLook(sh.cells[top]), ...look }
    if (a.r1 === a.r2 && a.c1 === a.c2) continue
    looks.push({ range: cornersText(a), ...look })
    for (const i of cellsOf(g, a)) {
      const addr = g.addr[i]
      if (addr === top) continue
      const kept = cleared(g.cell[i])
      if (kept) cells[addr] = { ...kept, ...look }
      else delete cells[addr]
    }
  }
  const out: SheetInput = {
    ...sh,
    cells,
    merges: [...kept, ...made.map(cornersText)],
    ...(looks.length ? { looks: [...(sh.looks ?? []), ...looks] } : {}),
  }
  if (!out.merges!.length) delete out.merges
  return { ok: true, book: { ...moved, sheets: moved.sheets.map((t, i) => (i === s ? out : t)), ...(styles.length ? { styles } : {}) }, geometry }
}

export function unmergeCells(book: BookInput, sheet: string, x: Rect): Applied {
  const s = book.sheets.findIndex((t) => t.name === sheet)
  const sh = book.sheets[s]
  const merges = (sh.merges ?? []).filter((m) => !meets(parseMerge(m)!, x))
  if (merges.length === (sh.merges ?? []).length) return { ok: true, book, geometry: [] }
  const out: SheetInput = { ...sh, merges }
  if (!merges.length) delete out.merges
  return { ok: true, book: { ...book, sheets: book.sheets.map((t, i) => (i === s ? out : t)) }, geometry: [] }
}

/** An entry without its style and number format. */
function withoutLook(cell: CellInput | undefined): CellInput {
  const { style: _s, numberFormat: _f, ...rest } = cell ?? {}
  return rest
}

/** What a cleared cell keeps (its meta, note and link), or null when nothing. */
function cleared(cell: CellInput): CellInput | null {
  const rest = withoutLook(cell)
  for (const k of CONTENT) delete rest[k]
  return Object.keys(rest).length ? rest : null
}
