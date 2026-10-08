/**
 * Tables under rows and columns inserted and deleted (the set `tables/ops`, Excel 16.113 for Mac). A table grows with
 * lines inserted after its first line and up to its last (the totals row included), moves with lines inserted at or
 * before its first, and keeps its size past its last; inserted rows get the formulas of its calculated columns,
 * inserted columns a free name `Column1`, `Column2` (`Столбец1` in a ru-RU book). Deleted lines shrink it; a table
 * left without data keeps one row, from below; deleting its totals row takes the row away; all its lines gone delete
 * it, and the references to it become #REF!, as do references to a deleted column (one alone or an end of a span).
 * Excel refuses lines inserted into two tables at once (a line inside both), deleted from two tables that keep some
 * of their lines, and a table's header row deleted without the rest.
 */
import type { BookInput, CellInput, SheetInput, TableInput } from '../contract.js'
import { formatAddress } from '../model.js'
import { sheetKey } from '../syntax.js'
import type { Axis } from '../structure/area.js'
import { rewriteBookTables } from './rewrite.js'
import { tableKey, tablesOf } from './model.js'

/** The word of default column names by the book's locale, as Excel names a new column. */
const COLUMN_WORD: Record<string, string> = { 'ru-RU': 'Столбец' }

/** The first free name `Column<n>` (n from 1) among `taken`, case-insensitively. */
export function freeColumnName(taken: string[], locale: string | undefined): string {
  const word = COLUMN_WORD[locale ?? ''] ?? 'Column'
  const keys = new Set(taken.map(tableKey))
  let n = 1
  while (keys.has(tableKey(`${word}${n}`))) n++
  return `${word}${n}`
}

const bounds = (t: { r1: number; r2: number; c1: number; c2: number }, axis: Axis): [number, number] => (axis === 'r' ? [t.r1, t.r2] : [t.c1, t.c2])

/** The effect of a shift on the tables of `sheet`: a refusal, or the book prepared before the shift and finished after it. */
export interface TableShift {
  refuse?: string
  /** References to deleted tables and columns made #REF!, before the cells move. */
  before(book: BookInput): BookInput
  /** The tables' new ranges and columns, the cells of inserted rows and columns, after the cells moved. */
  after(book: BookInput): BookInput
}

/** Lines inserted (`n` > 0) or deleted (`n` < 0) at index `at` of `axis` on `sheet`. */
export function tableShift(book: BookInput, sheet: string, axis: Axis, at: number, n: number): TableShift {
  const s = book.sheets.findIndex((x) => sheetKey(x.name) === sheetKey(sheet))
  const tables = tablesOf(book).list.filter((t) => t.s === s)
  const none: TableShift = { before: (b) => b, after: (b) => b }
  if (!tables.length) return none
  const count = Math.abs(n)
  const last = at + count - 1
  const lines = axis === 'r' ? 'rows' : 'columns'
  if (n > 0) {
    const inside = tables.filter((t) => bounds(t, axis)[0] <= at && at <= bounds(t, axis)[1])
    if (inside.length > 1) return { ...none, refuse: `inserting ${lines} there would shift cells of the tables ${inside.map((t) => t.name).join(' and ')} at once` }
  } else {
    const cut = tables.filter((t) => {
      const [lo, hi] = bounds(t, axis)
      return lo <= last && hi >= at && !(at <= lo && last >= hi)
    })
    if (cut.length > 1) return { ...none, refuse: `deleting these ${lines} would shift cells of the tables ${cut.map((t) => t.name).join(' and ')} at once` }
    const head = cut.find((t) => axis === 'r' && t.headers && t.r1 >= at && t.r1 <= last)
    if (head) return { ...none, refuse: `the header row of the table ${head.name} cannot be deleted without the rest of the table` }
  }
  // what each table becomes: its new range along the axis and its columns, or null when it is deleted
  const plan = new Map<string, { lo: number; hi: number; columns: TableInput['columns']; totals: boolean; grown: number[] } | null>()
  for (const t of tables) {
    const [lo, hi] = bounds(t, axis)
    const cols = t.input.columns
    if (n > 0) {
      if (at <= lo) plan.set(t.name, { lo: lo + count, hi: hi + count, columns: cols, totals: t.totals, grown: [] })
      else if (at <= hi) {
        const grown = Array.from({ length: count }, (_, i) => at + i)
        const columns = axis === 'c' ? [...cols.slice(0, at - t.c1), ...grown.map(() => ({ name: '' })), ...cols.slice(at - t.c1)] : cols
        plan.set(t.name, { lo, hi: hi + count, columns, totals: t.totals, grown })
      }
      continue
    }
    if (hi < at) continue
    if (lo > last) {
      plan.set(t.name, { lo: lo - count, hi: hi - count, columns: cols, totals: t.totals, grown: [] })
      continue
    }
    if (at <= lo && last >= hi) {
      plan.set(t.name, null)
      continue
    }
    const kept = (i: number) => i < at || i > last
    if (axis === 'c') {
      const columns = cols.filter((_, i) => kept(t.c1 + i))
      const nlo = lo < at ? lo : at
      plan.set(t.name, { lo: nlo, hi: nlo + columns.length - 1, columns, totals: t.totals, grown: [] })
      continue
    }
    const first = t.r1 + (t.headers ? 1 : 0)
    const lastData = t.r2 - (t.totals ? 1 : 0)
    let data = 0
    for (let r = first; r <= lastData; r++) if (kept(r)) data++
    const totals = t.totals && kept(t.r2)
    const nlo = lo < at ? lo : at
    plan.set(t.name, { lo: nlo, hi: nlo + (t.headers ? 1 : 0) + Math.max(1, data) + (totals ? 1 : 0) - 1, columns: cols, totals, grown: [] })
  }
  if (!plan.size) return none
  // a deleted column: references to it alone, or to an end of a span, are lost
  const gone = new Set(tables.filter((t) => plan.get(t.name) === null).map((t) => tableKey(t.name)))
  const deletedCols = axis === 'c' && n < 0 ? new Map(tables.map((t) => [tableKey(t.name), new Set(t.cols.filter((_, i) => t.c1 + i >= at && t.c1 + i <= last).map(tableKey))])) : null
  const before = (b: BookInput) =>
    gone.size || deletedCols
      ? rewriteBookTables(b, (t, spec) => {
          if (gone.has(tableKey(t.name))) return 'lost'
          const dead = deletedCols?.get(tableKey(t.name))
          if (!dead?.size || !spec.cols) return undefined
          const [a, z = a] = spec.cols
          return dead.has(tableKey(a)) || dead.has(tableKey(z)) ? 'lost' : undefined
        })
      : b
  const after = (b: BookInput): BookInput => {
    const sh = b.sheets[s]
    const cells = { ...sh.cells }
    const list: TableInput[] = []
    for (const ti of sh.tables ?? []) {
      const p = plan.get(ti.name)
      if (p === null) continue
      if (p === undefined) {
        list.push(ti)
        continue
      }
      const t = tables.find((x) => x.name === ti.name)!
      const r = axis === 'r' ? { r1: p.lo, r2: p.hi, c1: t.c1, c2: t.c2 } : { r1: t.r1, r2: t.r2, c1: p.lo, c2: p.hi }
      let columns = p.columns
      if (axis === 'c' && p.grown.length) {
        // new columns take the first free names, left to right
        columns = columns.map((c) => c)
        for (let i = 0; i < columns.length; i++) {
          if (columns[i].name !== '') continue
          columns[i] = { name: freeColumnName(columns.filter((c) => c.name !== '').map((c) => c.name), b.locale) }
          if (t.headers) cells[formatAddress(r.r1, r.c1 + i)] = { ...cells[formatAddress(r.r1, r.c1 + i)], value: columns[i].name }
        }
      }
      // inserted rows get the calculated columns' formulas
      if (axis === 'r')
        for (const row of p.grown)
          columns.forEach((c, i) => {
            if (c.formula === undefined) return
            const addr = formatAddress(row, r.c1 + i)
            cells[addr] = { ...cells[addr], formula: c.formula }
          })
      const next: TableInput = { ...ti, range: `${formatAddress(r.r1, r.c1)}:${formatAddress(r.r2, r.c2)}`, columns }
      if (p.totals) next.totals = true
      else delete next.totals
      list.push(next)
    }
    const out: SheetInput = { ...sh, cells }
    if (list.length) out.tables = list
    else delete out.tables
    const sheets = [...b.sheets]
    sheets[s] = out
    return { ...b, sheets }
  }
  return { before, after }
}
