/**
 * Excel's Sort (`sortRange`, and the sort `applyFilter` applies with a filter's criteria): the lines of a range shown
 * reordered by its conditions, ties keeping their order, the hidden ones staying where they are. A cell's content moves
 * as a copy lays it: relative references of its formula follow it, references from elsewhere keep their addresses (the
 * filter sets `sort-*`), so no cell of the book moves and `mapCell` keeps every address.
 */
import type { BookInput, CellInput, CellValue, FilterSort, SheetInput, SortCondition } from '../contract.js'
import type { EngineState } from '../engine.js'
import { InputError } from '../input-error.js'
import { checkSheetOf } from '../input.js'
import type { ValuesIn } from '../structure/values.js'
import { checkFilterSort } from './check.js'
import { entryAt, isHidden, lines } from '../layout/spans.js'
import { parseMerge } from '../layout/merges.js'
import { sayPivot } from '../locale/pivot-messages.js'
import { formatAddress, parseAddress } from '../model.js'
import { sameType, sortOrder } from '../sort-order.js'
import { contains, hasCell, meets, parseRange, rectText, type Rect } from '../structure/area.js'
import { relocated } from '../structure/arrays.js'
import type { Applied } from '../structure/geometry.js'
import { arraysOf } from '../structure/grid.js'
import { rewriteFormula, sheetNames } from '../structure/rewrite.js'
import { copyTransform } from '../structure/transform.js'
import { rowsOf, tablesOf } from '../tables/model.js'
import { areaCells, keepRules } from './cells.js'
import { KEEPS, textKey, type ColumnCells, type KeepRules } from './keep.js'

type Refused = Exclude<Applied, { ok: true }>

/** The order of two of a condition's lines, by their place in the range. */
type Order = (a: number, b: number) => number

/** `ß` as `ss`, which Excel's Sort takes as equal, either case taken into account or not (`sort-1-*`, `sort-2-*`, `sort-case`). */
const fold = (s: string) => s.replace(/ß/g, 'ss')

/** Text in Excel's case-sensitive order: as the locale orders it, then lower case before upper at the first letter that differs (`sort-2-True`). */
function caseOrder(compare: (a: string, b: string) => number): (a: string, b: string) => number {
  return (a, b) => {
    const x = compare(a, b)
    ;[a, b] = [fold(a), fold(b)]
    if (x || a === b) return x
    for (let i = 0; i < Math.min(a.length, b.length); i++) {
      if (a[i] === b[i]) continue
      const [la, lb] = [a[i] !== a[i].toUpperCase(), b[i] !== b[i].toUpperCase()]
      if (la !== lb) return la ? -1 : 1
    }
    return 0
  }
}

/** A condition's order of its line's cells: by value, a custom list's items first, or the cells of its color or icon first. */
function conditionOrder(cond: SortCondition, cells: ColumnCells, rules: KeepRules, caseSensitive: boolean): Order {
  const desc = cond.descending === true
  const by = cond.by ?? 'value'
  if (by !== 'value') {
    // the color or icon on top, the others after it as they were; descending, at the bottom
    const k =
      by === 'icon'
        ? cond.iconSet && KEEPS.icon({ type: 'icon', set: cond.iconSet, ...(cond.icon === undefined ? {} : { icon: cond.icon }) }, cells, rules)
        : KEEPS.color({ type: 'color', style: cond.style ?? {}, ...(by === 'fontColor' ? { font: true } : {}) }, cells, rules)
    const keep = k && 'keep' in k ? k.keep : () => false
    return (a, b) => (Number(keep(b)) - Number(keep(a))) * (desc ? -1 : 1)
  }
  const folded = (a: string, b: string) => rules.compareText(fold(a), fold(b))
  const text = caseSensitive ? caseOrder(folded) : folded
  const same = sameType(text) as (a: CellValue, b: CellValue) => number
  const values = new Map<number, CellValue>()
  const value = (i: number) => (values.has(i) ? values.get(i)! : (values.set(i, cells.value(i)), values.get(i)!))
  if (cond.customList === undefined) return (a, b) => sortOrder(value(a), value(b), desc, same)
  // a custom list's items in its order, then the other values in theirs, blanks last (`sort-custom`)
  const items = new Map(cond.customList.split(',').map((x, i) => [textKey(x.trim()), i]))
  const place = (i: number) => (value(i) === null ? undefined : items.get(textKey(cells.text(i).trim())))
  return (a, b) => {
    const [x, y] = [place(a), place(b)]
    if (x !== undefined || y !== undefined) return x === undefined ? 1 : y === undefined ? -1 : desc ? y - x : x - y
    return sortOrder(value(a), value(b), desc, same)
  }
}

/** The line a condition sorts by, inside the range: its column, or its row with `byColumns`. */
function keyLine(cond: SortCondition, area: Rect, byColumns: boolean, path: string): number {
  const x = parseRange(cond.range)!
  const at = byColumns ? x.r1 : x.c1
  const inside = byColumns ? at >= area.r1 && at <= area.r2 : at >= area.c1 && at <= area.c2
  if (!inside) throw new InputError(`${path}.range`, `the sort condition ${cond.range} is outside the sorted range`)
  return at
}

/**
 * Why Excel refuses to sort `area` of sheet `s`, or null: a merged area it meets, a legacy array formula it meets other
 * than whole in one of its lines, a table it meets without lying in the table's data.
 */
function sortRefusal(book: BookInput, s: number, area: Rect, byColumns: boolean): Refused | null {
  const sheet = book.sheets[s]
  const no = (code: Refused['code'], message: Parameters<typeof sayPivot>[1]): Refused => ({ ok: false, code, message: sayPivot(book.locale, message) })
  const merged = (sheet.merges ?? []).find((m) => meets(parseMerge(m)!, area))
  if (merged) return no('merged', { key: 'sortMerged', sheet: sheet.name, area: merged })
  for (const [addr, x] of arraysOf(sheet)) {
    const oneLine = byColumns ? x.c1 === x.c2 : x.r1 === x.r2
    if (meets(x, area) && (!contains(area, x) || !oneLine)) return no('array', { key: 'arrayPart', sheet: sheet.name, cell: addr, area: rectText(x) })
  }
  for (const t of tablesOf(book).list) {
    if (t.s !== s || !meets(t, area)) continue
    const { first, last } = rowsOf(t)
    if (!contains({ r1: first, r2: last, c1: t.c1, c2: t.c2 }, area)) return no('table', { key: 'sortTable', table: t.name })
  }
  return null
}

/**
 * The book with `area` of sheet `s` sorted by `sort` as the state computed it: the lines `hidden` says stay; null when no
 * line moves. Conditions outside the range are an input error at `path`.
 */
export function sortedBook(book: BookInput, state: EngineState, s: number, area: Rect, sort: FilterSort, hidden: (line: number) => boolean, path: string): { ok: true; book: BookInput } | Refused | null {
  const byColumns = sort.byColumns === true
  const [lo, hi] = byColumns ? [area.c1, area.c2] : [area.r1, area.r2]
  const conds = sort.conditions ?? []
  const keys = conds.map((c, i) => keyLine(c, area, byColumns, `${path}.conditions[${i}]`))
  const shown: number[] = []
  for (let k = lo; k <= hi; k++) if (!hidden(k)) shown.push(k)
  if (!conds.length) return null
  const refused = sortRefusal(book, s, area, byColumns)
  if (refused) return refused
  if (shown.length < 2) return null
  const rules = keepRules(state)
  const cellsOf = areaCells(state, s, area)
  const n = hi - lo + 1
  const orders = conds.map((c, i) => conditionOrder(c, cellsOf(n, (j) => (byColumns ? [keys[i], lo + j] : [lo + j, keys[i]])), rules, sort.caseSensitive === true))
  const order = shown.slice().sort((a, b) => {
    for (const o of orders) {
      const x = o(a - lo, b - lo)
      if (x) return x
    }
    return a - b
  })
  if (order.every((x, k) => x === shown[k])) return null
  // the line each shown line's content comes from
  const from = new Map(shown.map((to, k) => [order[k], to]))
  const sheet = book.sheets[s]
  const names = sheetNames(book)
  const cells: Record<string, CellInput> = {}
  const moved: [string, CellInput, number][] = []
  for (const [addr, cell] of Object.entries(sheet.cells)) {
    const { row, col } = parseAddress(addr)
    const line = byColumns ? col : row
    const to = hasCell(area, row, col) ? from.get(line) : undefined
    // a line whose content leaves gets another's, or none
    if (to === undefined || to === line) cells[addr] = cell
    else moved.push([addr, cell, to - line])
  }
  for (const [addr, cell, d] of moved) {
    const { row, col } = parseAddress(addr)
    const dest = byColumns ? formatAddress(row, col + d) : formatAddress(row + d, col)
    const formula = cell.formula === undefined ? undefined : rewriteFormula(cell.formula, sheet.name, sheet.name, byColumns ? copyTransform(0, d) : copyTransform(d, 0), names)
    cells[dest] = relocated(formula === cell.formula ? cell : { ...cell, formula }, addr, dest)
  }
  const sheets = book.sheets.slice()
  sheets[s] = { ...sheet, cells }
  return { ok: true, book: { ...book, sheets } }
}

/** Whether a line of the sheet is hidden, a row or a column, of height or width 0 too. */
export function hiddenLine(sheet: SheetInput, byColumns: boolean): (line: number) => boolean {
  const ls = lines<{ hidden?: boolean; height?: number; width?: number }>(byColumns ? sheet.columns : sheet.rows, byColumns ? 'c' : 'r')
  return (k) => {
    const e = entryAt(ls, k)
    return !!e && isHidden(e)
  }
}

/** `sortRange` on a book already checked; its fields but `op` are checked here. */
export function sortOp(book: BookInput, o: Record<string, unknown>, values: ValuesIn): Applied {
  const sheet = checkSheetOf(book, o.sheet, 'op.sheet')
  const s = book.sheets.indexOf(sheet)
  checkFilterSort(o.sort, 'op.sort')
  const sort = JSON.parse(JSON.stringify(o.sort)) as FilterSort
  const area = parseRange(sort.range)!
  const sorted = sortedBook(book, values.state!(), s, area, sort, hiddenLine(sheet, sort.byColumns === true), 'op.sort')
  if (sorted && !sorted.ok) return sorted
  const out = sorted?.book ?? book
  const sheets = out.sheets.slice()
  // kept as the sheet's, over a filter's data too, as `Worksheet.Sort` keeps it (`sort-in-filter`, `sort-in-table`)
  sheets[s] = { ...sheets[s], sort }
  return { ok: true, book: { ...out, sheets }, geometry: [] }
}
