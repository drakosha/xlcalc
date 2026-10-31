/** A PivotTable's cells on its sheet: its range and, above it, its Filters area; its areas as a host reads them. */
import type { BookInput, PivotAreas, PivotTable, SheetInput } from '../contract.js'
import { parseRange, rectText, type Rect } from '../structure/area.js'

/** The width of the Filters area past its first column: each line of filters two cells wide and one apart (`pageColumns`). */
const pageWidth = (t: Pick<PivotTable, 'location'>) => 3 * Math.max(1, t.location.pageColumns ?? 1) - 2

/**
 * The rectangle of a PivotTable's cells: its range, and the rows of its Filters area above it with the blank row between
 * (`pageRows` + 1), each line of filters two cells wide and one apart (`pageColumns`); null for a range that does not parse.
 */
export function pivotArea(t: Pick<PivotTable, 'location'>): Rect | null {
  const x = parseRange(t.location.range)
  if (!x) return null
  const rows = t.location.pageRows ?? 0
  if (!rows) return x
  return { r1: Math.max(0, x.r1 - rows - 1), c1: x.c1, r2: x.r2, c2: Math.max(x.c2, x.c1 + pageWidth(t)) }
}

/** The Filters area as wide as its filters, without the blank row under it; null without one. */
function pageArea(t: Pick<PivotTable, 'location'>, x: Rect): Rect | null {
  const rows = t.location.pageRows ?? 0
  const r1 = Math.max(0, x.r1 - rows - 1)
  return rows && r1 <= x.r1 - 2 ? { r1, r2: x.r1 - 2, c1: x.c1, c2: x.c1 + pageWidth(t) } : null
}

/**
 * The parts of a PivotTable an entry cannot change and a refresh clears: its range and its Filters area as wide as its
 * filters, not the blank row between them nor the cells beside the filters, which hold anything as in Excel.
 */
export function pivotParts(t: Pick<PivotTable, 'location'>): Rect[] {
  const x = parseRange(t.location.range)
  if (!x) return []
  const page = pageArea(t, x)
  return page ? [x, page] : [x]
}

type Placed = Pick<PivotTable, 'location' | 'rows' | 'columns' | 'values' | 'rowLines' | 'columnLines'>

/**
 * A PivotTable's areas (`PivotLocation.areas`) from its place and saved layout: the table and its Filters area; once laid
 * out, the header rows above its values, the labels of its rows left of them, and its values when it has value fields.
 * One not laid out (no saved lines), shrunk to its top-left cell (`#SPILL!`) or with no field at all has its table alone.
 */
export function pivotAreas(t: Placed): PivotAreas {
  const { range, firstDataRow: dr, firstDataColumn: dc } = t.location
  const x = parseRange(range)
  const out: PivotAreas = { table: range }
  if (!x) return out
  const page = pageArea(t, x)
  if (page) out.page = rectText(page)
  const empty = !t.rows?.length && !t.columns?.length && !t.values?.length
  const top = x.r1 + dr
  if (!t.rowLines || dr <= 0 || top > x.r2 || empty) return out
  const left = x.c1 + dc
  out.columnHeaders = rectText({ ...x, r2: top - 1 })
  if (dc > 0) out.rowLabels = rectText({ ...x, r1: top, c2: Math.min(x.c2, left - 1) })
  if (t.values?.length && left <= x.c2) out.data = rectText({ ...x, r1: top, c1: left })
  return out
}

const sameAreas = (a: PivotAreas | undefined, b: PivotAreas) => a !== undefined && a.table === b.table && a.page === b.page && a.columnHeaders === b.columnHeaders && a.rowLabels === b.rowLabels && a.data === b.data && Object.keys(a).length === Object.keys(b).length

/** The PivotTable with its areas worked out again from its place; the same object when they are already so. */
export function withAreas<T extends Placed>(t: T): T {
  const areas = pivotAreas(t)
  return sameAreas(t.location.areas, areas) ? t : { ...t, location: { ...t.location, areas } }
}

/** The sheet with its PivotTables' areas worked out again; the same object when none changes. */
const placedSheet = (s: SheetInput): SheetInput => {
  const list = s.pivotTables?.map(withAreas)
  return list && list.some((t, k) => t !== s.pivotTables![k]) ? { ...s, pivotTables: list } : s
}

/**
 * The book with the areas of its PivotTables worked out from their places, as the library gives every PivotTable out
 * and takes none from its input; the same object when none changes. `before` names the book whose sheets are known to
 * be so already: a sheet whose PivotTables are its own is not looked at.
 */
export function placedPivots<B extends BookInput>(book: B, before?: BookInput): B {
  let sheets: SheetInput[] | undefined
  book.sheets.forEach((s, i) => {
    if (!s.pivotTables || (before && before.sheets.some((o) => o.pivotTables === s.pivotTables))) return
    const placed = placedSheet(s)
    if (placed !== s) (sheets ??= book.sheets.slice())[i] = placed
  })
  return sheets ? { ...book, sheets } : book
}
