/**
 * Checks of a book's tables at the API boundary (`InputError`): names by Excel's rules in the namespace of defined names,
 * a rectangle with a data row, one column entry per column with unique names, header cells holding the names, no
 * table over another, no merged area or array area of several cells in a table.
 */
import type { BookInput, CellInput, SheetInput, TableColumn, TableInput, TableStyle, TotalsFunction } from '../contract.js'
import { AreaIndex } from '../area-index.js'
import type { Area } from '../spill/state.js'
import { InputError } from '../input-error.js'
import { isObject, onlyKeys } from '../input/basic.js'
import { parseMerge } from '../layout/merges.js'
import { formatAddress } from '../model.js'
import { arrayArea } from '../array-areas.js'
import { parseRange, type Rect } from '../structure/area.js'
import { localeBooleans } from '../locale/registry.js'
import { nameProblem } from '../syntax.js'
import { tableKey } from './model.js'

/** The functions of a totals row (`TotalsFunction`), as xlsx names them. */
export const TOTALS: readonly TotalsFunction[] = ['sum', 'average', 'count', 'countNums', 'max', 'min', 'stdDev', 'var', 'custom']

/** Excel's table styles by name: Light 1 to 21, Medium 1 to 28, Dark 1 to 11. */
export const TABLE_STYLE = /^TableStyle(?:Light(?:[1-9]|1\d|2[01])|Medium(?:[1-9]|1\d|2[0-8])|Dark(?:[1-9]|1[01]))$/

/** The options of a table's style (`TableStyle`), each with its name in xlsx `tableStyleInfo`. */
export const STYLE_OPTIONS: readonly [Exclude<keyof TableStyle, 'name'>, string][] = [
  ['firstColumn', 'showFirstColumn'],
  ['lastColumn', 'showLastColumn'],
  ['rowStripes', 'showRowStripes'],
  ['columnStripes', 'showColumnStripes'],
]

/** Excel's default style, a table's without `style`. */
export const DEFAULT_TABLE_STYLE: TableStyle = { name: 'TableStyleMedium2', rowStripes: true }

function style(v: unknown, path: string) {
  if (!isObject(v)) throw new InputError(path, 'expected a style such as { name: "TableStyleLight1", rowStripes: true }')
  onlyKeys(v, path, ['name', ...STYLE_OPTIONS.map(([k]) => k)])
  if (v.name !== undefined && (typeof v.name !== 'string' || !TABLE_STYLE.test(v.name)))
    throw new InputError(`${path}.name`, `expected one of Excel's table styles, TableStyleLight1 to 21, TableStyleMedium1 to 28 or TableStyleDark1 to 11, got ${JSON.stringify(v.name)}`)
  for (const [k] of STYLE_OPTIONS) if (v[k] !== undefined && typeof v[k] !== 'boolean') throw new InputError(`${path}.${k}`, 'expected a boolean')
}

function column(v: unknown, path: string): TableColumn {
  if (!isObject(v)) throw new InputError(path, 'expected a column such as { name: "Amount" }')
  onlyKeys(v, path, ['name', 'formula', 'totals'])
  if (typeof v.name !== 'string' || v.name === '' || v.name.length > 255) throw new InputError(`${path}.name`, 'expected a column name of 1 to 255 characters')
  if (v.formula !== undefined && typeof v.formula !== 'string') throw new InputError(`${path}.formula`, 'expected formula text')
  if (v.totals !== undefined) {
    const t = v.totals
    const ok = isObject(t) && Object.keys(t).length === 1 && (TOTALS.includes(t.function as TotalsFunction) || typeof t.label === 'string')
    if (!ok) throw new InputError(`${path}.totals`, `expected { function } with one of ${TOTALS.join(', ')}, or { label }`)
  }
  return v as unknown as TableColumn
}

/** One table's own shape: its fields, its rectangle and columns, its header cells; returns the rectangle. */
function table(v: unknown, path: string, cellAt: (addr: string) => CellInput | undefined, booleans: readonly string[]): Rect {
  if (!isObject(v)) throw new InputError(path, 'expected a table such as { name: "Sales", range: "A1:C9", columns: [...] }')
  onlyKeys(v, path, ['name', 'range', 'headers', 'totals', 'columns', 'style'])
  if (typeof v.name !== 'string') throw new InputError(`${path}.name`, 'expected a table name')
  const problem = nameProblem(v.name, booleans)
  if (problem) throw new InputError(`${path}.name`, `${JSON.stringify(v.name)}: ${problem}`)
  const x = typeof v.range === 'string' && /^[A-Z]+\d+(:[A-Z]+\d+)?$/.test(v.range) ? parseRange(v.range) : null
  if (!x) throw new InputError(`${path}.range`, `expected a range such as "B2:E10" within the sheet, got ${JSON.stringify(v.range)}`)
  for (const k of ['headers', 'totals'] as const) if (v[k] !== undefined && typeof v[k] !== 'boolean') throw new InputError(`${path}.${k}`, 'expected a boolean')
  if (v.style !== undefined) style(v.style, `${path}.style`)
  const t = v as unknown as TableInput
  const headers = t.headers !== false
  if (x.r2 - x.r1 + 1 < (headers ? 1 : 0) + (t.totals ? 1 : 0) + 1) throw new InputError(`${path}.range`, `${t.range} has no row for data between its header and totals rows`)
  if (!Array.isArray(v.columns)) throw new InputError(`${path}.columns`, 'expected an array of columns')
  const width = x.c2 - x.c1 + 1
  if (v.columns.length !== width) throw new InputError(`${path}.columns`, `${t.range} has ${width} columns, but the table lists ${v.columns.length}`)
  const seen = new Map<string, number>()
  v.columns.forEach((c, i) => {
    const col = column(c, `${path}.columns[${i}]`)
    const k = tableKey(col.name)
    if (seen.has(k)) throw new InputError(`${path}.columns[${i}].name`, `${JSON.stringify(col.name)} is also the name of column ${seen.get(k)! + 1} (column names are unique in a table, regardless of case)`)
    seen.set(k, i)
    if (!headers) return
    const addr = formatAddress(x.r1, x.c1 + i)
    const cell = cellAt(addr)
    if (cell?.formula !== undefined || cell?.value !== col.name)
      throw new InputError(`${path}.columns[${i}].name`, `the header cell ${addr} holds ${cell?.formula !== undefined ? `a formula` : JSON.stringify(cell?.value ?? null)}, not the column's name ${JSON.stringify(col.name)} as text`)
  })
  return x
}

/** A rectangle among those a sheet's tables are checked against, with its text and its place in its list. */
interface Box extends Area {
  text: string
  i: number
}

/** The boxes of rectangles in an index (one sheet, `s` 0), each by its place: overlaps found by the area, not by a scan of all. */
function indexed(xs: [Rect, string][]): AreaIndex<Box> {
  const index = new AreaIndex<Box>()
  xs.forEach(([x, text], i) => index.add({ ...x, s: 0, text, i }))
  return index
}

/** The first box, by its place, that meets `x`. */
function firstOver(index: AreaIndex<Box>, x: Rect): Box | undefined {
  let first: Box | undefined
  index.forEach({ ...x, s: 0 }, (b) => {
    if (!first || b.i < first.i) first = b
  })
  return first
}

/** A sheet's array formulas of several cells: each area with its text, in the order of the sheet's cells. */
function wideArrays(cells: Record<string, CellInput>): [Rect, string][] {
  const out: [Rect, string][] = []
  for (const [addr, cell] of Object.entries(cells)) {
    if (cell.array === undefined) continue
    const a = arrayArea(addr, cell.array)
    if (a && (a.r2 > a.r1 || a.c2 > a.c1)) out.push([a, cell.array])
  }
  return out
}

/** Why an AutoFilter cannot lie over a table's cells. */
export const autoFilterOverlap = (range: string, table: string) => `the AutoFilter over ${range} overlaps the table ${table} (a table has a filter of its own)`

/**
 * A sheet's tables checked one by one (`add`), each against the book's defined names (`names`) and table names
 * (`taken`, which it joins) and the sheet's tables before it, merged areas and array formulas, by an index of their
 * rectangles: a sheet of many tables costs about as many checks as it has tables. A table that fails changes nothing.
 */
export function sheetTables(sheet: Pick<SheetInput, 'cells' | 'merges'>, booleans: readonly string[], names: Map<string, string>, taken: Map<string, string>) {
  const rects = new AreaIndex<Box>()
  let count = 0
  // indexed once for all the sheet's tables, where the first table gets to them
  let arrays: AreaIndex<Box> | undefined
  let merges: AreaIndex<Box> | undefined
  return {
    /** The table `t` at `tp`, its header cells read through `cellAt`; its rectangle. */
    add(t: unknown, tp: string, cellAt: (addr: string) => CellInput | undefined): Rect {
      const x = table(t, tp, cellAt, booleans)
      const { name, range } = t as TableInput
      const k = tableKey(name)
      if (taken.has(k)) throw new InputError(`${tp}.name`, `${JSON.stringify(name)} is also the name of the table ${taken.get(k)} (table names are unique in a workbook, regardless of case)`)
      if (names.has(k)) throw new InputError(`${tp}.name`, `${JSON.stringify(name)} is also a defined name (tables and names share one namespace, regardless of case)`)
      const over = firstOver(rects, x)
      if (over) throw new InputError(`${tp}.range`, `${range} overlaps the table ${over.text}`)
      const merge = firstOver((merges ??= indexed((sheet.merges ?? []).map((m) => [parseMerge(m)!, m]))), x)
      if (merge) throw new InputError(`${tp}.range`, `the merged area ${merge.text} overlaps the table ${name} (Excel does not allow merged cells in tables)`)
      const array = firstOver((arrays ??= indexed(wideArrays(sheet.cells))), x)
      if (array) throw new InputError(`${tp}.range`, `the array formula over ${array.text} overlaps the table ${name} (Excel does not allow array formulas of several cells in tables)`)
      taken.set(k, name)
      rects.add({ ...x, s: 0, text: name, i: count++ })
      return x
    },
    /** The first table, by its place, over `x`: its name. */
    over: (x: Rect) => firstOver(rects, x)?.text,
  }
}

/** The tables of a checked book's sheets, and their names against each other and against the defined names; no sheet's AutoFilter over one. */
export function checkTables(book: BookInput, path = 'book'): void {
  const names = new Map((book.names ?? []).map((n) => [tableKey(n.name), n.name]))
  const taken = new Map<string, string>()
  const booleans = localeBooleans(book.locale)
  book.sheets.forEach((sheet, s) => {
    if (sheet.tables === undefined) return
    const p = `${path}.sheets[${s}].tables`
    if (!Array.isArray(sheet.tables)) throw new InputError(p, 'expected an array of tables')
    const tables = sheetTables(sheet, booleans, names, taken)
    sheet.tables.forEach((t, i) => tables.add(t, `${p}[${i}]`, (addr) => sheet.cells[addr]))
    const f = sheet.autoFilter && tables.over(parseRange(sheet.autoFilter.range)!)
    if (f) throw new InputError(`${path}.sheets[${s}].autoFilter.range`, autoFilterOverlap(sheet.autoFilter!.range, f))
  })
}
