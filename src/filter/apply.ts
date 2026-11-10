/**
 * Excel's filter commands as structure operations (`applyFilter`, `clearFilter`): the criteria of a sheet's AutoFilter
 * or a table's set or cleared, and applied to the rows of its range, which show or hide by them as Excel's Reapply shows
 * and hides them; no cell moves.
 */
import type { BookInput, ColumnFilter, FilterColumn, FilterCriteria, RowInput, SheetInput, StructureOp, TableInput } from '../contract.js'
import { sayPivot } from '../locale/pivot-messages.js'
import { checkSheetOf } from '../input.js'
import { InputError } from '../input-error.js'
import { checkString, integer } from '../input/basic.js'
import { lines, toMap, type Line } from '../layout/spans.js'
import { formatAddress } from '../model.js'
import { autoFilterRange } from '../session/autofilter.js'
import { parseRange, rectText, type Rect } from '../structure/area.js'
import type { Applied } from '../structure/geometry.js'
import type { ValuesIn } from '../structure/values.js'
import { rowsOf, tableKey, tablesOf } from '../tables/model.js'
import { checkFilterCriteria } from './check.js'
import { areaCells, keepRules } from './cells.js'
import { KEEPS, type ColumnCells, type KeepRules, type Kept } from './keep.js'
import { hiddenLine, sortedBook } from './sort.js'
import { sheetKey } from '../syntax.js'

export type FilterOp = Extract<StructureOp, { op: 'applyFilter' | 'clearFilter' }>

/** The fields each filter operation takes beside `op`. */
export const FILTER_FIELDS: Record<FilterOp['op'], string[]> = {
  applyFilter: ['sheet', 'table', 'criteria'],
  clearFilter: ['sheet', 'table', 'column'],
}

/** A filter of a sheet: its sheet, its data rows and columns, its criteria, and the book with new ones. */
interface Target {
  s: number
  data: Rect
  criteria: FilterCriteria
  put(sheet: SheetInput, criteria: FilterCriteria): SheetInput
}

/** A formula that is itself a SUBTOTAL, whatever its arguments: `SUBTOTAL(9,C2:C21)`, not `ROUND(SUBTOTAL(…),0)`. */
function isSubtotal(formula: string | undefined): boolean {
  if (!formula || !/^SUBTOTAL\(/i.test(formula)) return false
  let depth = 0
  for (let i = 8; i < formula.length; i++) {
    const ch = formula[i]
    if (ch === '"' || ch === "'") {
      const end = formula.indexOf(ch, i + 1)
      if (end < 0) return false
      i = end
    } else if (ch === '(') depth++
    else if (ch === ')' && --depth === 0) return i === formula.length - 1
  }
  return false
}

/** Criteria with something to hold: a column or a sort. */
const held = (c: FilterCriteria) => !!c.columns?.length || c.sort !== undefined

/** Without the criteria, a table's filter is `{}`: it has its buttons. */
const tableCriteria = (t: TableInput): FilterCriteria => (t.filter === false ? {} : (t.filter ?? {}))

function targetOf(book: BookInput, o: Record<string, unknown>, values: ValuesIn): Target {
  if ((o.sheet === undefined) === (o.table === undefined)) throw new InputError('op', 'expected one of sheet and table')
  if (o.table !== undefined) {
    const name = checkString(o.table, 'op.table')
    const t = tablesOf(book).byName.get(tableKey(name))
    if (!t) throw new InputError('op.table', `unknown table ${JSON.stringify(name)}`)
    if (!t.headers || t.input.filter === false) throw new InputError('op.table', `the table ${t.name} has no filter buttons`)
    const { first, last } = rowsOf(t)
    return {
      s: t.s,
      data: { r1: first, r2: last, c1: t.c1, c2: t.c2 },
      criteria: tableCriteria(t.input),
      put: (sheet, criteria) => ({
        ...sheet,
        tables: sheet.tables!.map((x) => {
          if (tableKey(x.name) !== tableKey(t.name)) return x
          const { filter: _, ...rest } = x
          return held(criteria) ? { ...rest, filter: criteria } : rest
        }),
      }),
    }
  }
  const sheet = checkSheetOf(book, o.sheet, 'op.sheet')
  const s = book.sheets.indexOf(sheet)
  if (!sheet.autoFilter) throw new InputError('op.sheet', `the sheet ${sheet.name} has no AutoFilter`)
  const state = values.state!()
  // the range Excel reads, without a last row of SUBTOTAL, which it neither filters nor saves (docs/structure.md)
  const read = parseRange(autoFilterRange(state, s)!)!
  let r2 = read.r2
  if (r2 > read.r1) {
    const cells = state.input.sheets[s].cells
    for (let c = read.c1; c <= read.c2; c++) {
      if (!isSubtotal(cells[formatAddress(r2, c)]?.formula)) continue
      r2--
      break
    }
  }
  const range = rectText({ ...read, r2 })
  return {
    s,
    data: { r1: read.r1 + 1, r2, c1: read.c1, c2: read.c2 },
    criteria: sheet.autoFilter.criteria ?? {},
    put: (x, criteria) => ({ ...x, autoFilter: { range, ...(held(criteria) ? { criteria } : {}) } }),
  }
}

/**
 * The range `applyFilter` would sort, by the sheet it is on, read before the operation is checked (`structure/looks.ts`
 * lays its looks as entries): the sort of the criteria given, or else of the filter's own; none for input of another shape.
 */
export function filterSorts(book: BookInput, o: Record<string, unknown>): [number, Rect][] {
  let s = -1
  let held: FilterCriteria | false | undefined
  if (typeof o.table === 'string') {
    const t = tablesOf(book).byName.get(tableKey(o.table))
    if (t) [s, held] = [t.s, t.input.filter]
  } else if (typeof o.sheet === 'string') {
    s = book.sheets.findIndex((x) => sheetKey(x.name) === sheetKey(o.sheet as string))
    held = book.sheets[s]?.autoFilter?.criteria
  }
  const given = o.criteria as FilterCriteria | undefined
  const range = (given !== undefined ? (typeof given === 'object' && given ? given.sort?.range : undefined) : held ? held.sort?.range : undefined) as unknown
  const x = s >= 0 && typeof range === 'string' ? parseRange(range) : null
  return x ? [[s, x]] : []
}

/** The criteria without the conditions of their columns, or of column `only`: what is left of a column is its buttons. */
function cleared(criteria: FilterCriteria, only?: number): FilterCriteria {
  const columns = (criteria.columns ?? []).flatMap((x): FilterColumn[] => {
    if (only !== undefined && x.column !== only) return [x]
    const { filter: _, ...rest } = x
    return Object.keys(rest).length > 1 ? [rest] : []
  })
  const { columns: _, ...rest } = criteria
  return columns.length ? { ...rest, columns } : rest
}

/** The entry of a row hidden or shown, the same object when it does not change. */
const HIDDEN: RowInput = { hidden: true }
function rowEntry(v: RowInput | undefined, hide: boolean): RowInput | undefined {
  if (!!v?.hidden === hide) return v
  if (!v) return HIDDEN
  const { hidden: _, ...rest } = v
  return hide ? { ...rest, hidden: true } : Object.keys(rest).length ? rest : undefined
}

/** A sheet's rows with rows `first` to `last` hidden where `hide` says, in one pass over them; the others as they are. */
function hiddenRows(ls: Line<RowInput>[], first: number, last: number, hide: (r: number) => boolean): Line<RowInput>[] {
  const out: Line<RowInput>[] = []
  const push = (a: number, b: number, v: RowInput | undefined) => {
    if (!v) return
    const end = out[out.length - 1]
    if (end && end[1] === a - 1 && end[2] === v) end[1] = b
    else out.push([a, b, v])
  }
  let j = 0
  for (; j < ls.length && ls[j][1] < first; j++) out.push(ls[j])
  if (j < ls.length && ls[j][0] < first) out.push([ls[j][0], first - 1, ls[j][2]])
  for (let r = first; r <= last; ) {
    while (j < ls.length && ls[j][1] < r) j++
    const line = j < ls.length && ls[j][0] <= r ? ls[j] : undefined
    const end = Math.min(last, line ? line[1] : j < ls.length ? ls[j][0] - 1 : last)
    for (let a = r; a <= end; ) {
      const h = hide(a)
      let b = a
      while (b < end && hide(b + 1) === h) b++
      push(a, b, rowEntry(line?.[2], h))
      a = b + 1
    }
    r = end + 1
  }
  for (; j < ls.length; j++) if (ls[j][1] > last) push(Math.max(ls[j][0], last + 1), ls[j][1], ls[j][2])
  return out
}

/** Whether criteria hold a condition: a sheet with one is filtered. */
const conditioned = (c: FilterCriteria | false | undefined) => !!c && !!c.columns?.some((x) => x.filter)

/** The operation on a book already checked; its fields but `op` are checked here. */
export function filterOp(book: BookInput, o: Record<string, unknown>, values: ValuesIn): Applied {
  const t = targetOf(book, o, values)
  const width = t.data.c2 - t.data.c1 + 1
  let criteria = t.criteria
  if (o.op === 'applyFilter' && o.criteria !== undefined) {
    checkFilterCriteria(o.criteria, 'op.criteria', width)
    criteria = JSON.parse(JSON.stringify(o.criteria)) as FilterCriteria
  }
  if (o.op === 'clearFilter') criteria = cleared(criteria, o.column === undefined ? undefined : integer(o.column, 'op.column', 0, width - 1))
  const state = values.state!()
  const rules = keepRules(state)
  const n = Math.max(0, t.data.r2 - t.data.r1 + 1)
  const cellsOf = areaCells(state, t.s, t.data)
  const kept: ((i: number) => boolean)[] = []
  const columns: FilterColumn[] = []
  for (const x of criteria.columns ?? []) {
    const f = x.filter
    if (!f) {
      columns.push(x)
      continue
    }
    const c = t.data.c1 + x.column
    const row = (i: number) => t.data.r1 + i
    const cells = cellsOf(n, (i) => [row(i), c])
    const k = (KEEPS[f.type] as (f: ColumnFilter, cells: ColumnCells, rules: KeepRules) => Kept)(f, cells, rules)
    if ('error' in k) return { ok: false, code: 'filter', message: sayPivot(book.locale, { key: 'filterError', sheet: book.sheets[t.s].name, cell: formatAddress(row(k.error), c) }) }
    kept.push(k.keep)
    columns.push(k.filter ? { ...x, filter: k.filter } : x)
  }
  if (criteria.columns) criteria = { ...criteria, columns }
  const sheet = book.sheets[t.s]
  const rows = n ? toMap(hiddenRows(lines(sheet.rows, 'r'), t.data.r1, t.data.r2, (r) => !kept.every((keep) => keep(r - t.data.r1))), 'r') : sheet.rows
  const next: SheetInput = t.put(sheet, criteria)
  if (rows && Object.keys(rows).length) next.rows = rows
  else delete next.rows
  if (conditioned(next.autoFilter?.criteria) || next.tables?.some((x) => conditioned(x.filter))) next.filtered = true
  else delete next.filtered
  const sheets = book.sheets.slice()
  sheets[t.s] = next
  const out = { ...book, sheets }
  // Reapply sorts the rows it shows by the filter's sort too (`file-sort-stale`)
  if (o.op !== 'applyFilter' || !criteria.sort?.conditions?.length) return { ok: true, book: out, geometry: [] }
  const sorted = sortedBook(out, state, t.s, parseRange(criteria.sort.range)!, criteria.sort, hiddenLine(next, criteria.sort.byColumns === true), 'op.criteria.sort')
  return sorted && !sorted.ok ? sorted : { ok: true, book: sorted?.book ?? out, geometry: [] }
}
