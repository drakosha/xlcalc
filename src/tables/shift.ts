/**
 * Tables under rows and columns inserted and deleted (the set `tables/ops`, Excel 16.113 for Mac). A table grows with
 * lines inserted after its first line and up to its last (the totals row included), moves with lines inserted at or
 * before its first, and keeps its size past its last; inserted rows get the formulas of its calculated columns,
 * inserted columns a free name `Column1`, `Column2` (`Столбец1` in a ru-RU book). Deleted lines shrink it; a table
 * left without data keeps one row, from below; deleting its totals row takes the row away; all its lines gone delete
 * it, and the references to it become #REF!, as do references to a deleted column (one alone or an end of a span).
 * Excel refuses lines inserted into two tables at once (a line inside both), deleted from two tables when one of them
 * keeps some of its lines, and a table's header row deleted without the rest.
 */
import type { BookInput, SheetInput, TableInput } from '../contract.js'
import type { EngineMessage } from '../locale/messages.js'
import { localeOf } from '../locale/registry.js'
import { formatAddress } from '../model.js'
import { cornersText, hi, LINES, LIMIT, lo, rectOf, type Area, type Axis } from '../structure/area.js'
import { copyTransform } from '../structure/transform.js'
import { rewriteFormula, sheetIndex, sheetNames } from '../structure/rewrite.js'
import { rewriteBookTables } from './rewrite.js'
import { withField } from '../looks/sheet.js'
import { numbered, rowsOf, tableKey, tablesOf } from './model.js'

/** The first free name `Column<n>` (n from 1) among `taken`, case-insensitively, with the word of the book's locale (`columnWord`). */
export function freeColumnName(taken: string[], locale: string | undefined): string {
  const word = localeOf(locale).regional?.columnWord ?? 'Column'
  return numbered(word, taken, 1)
}

/** The effect of a shift on the tables of `sheet`: a refusal, or the book prepared before the shift and finished after it. */
export interface TableShift {
  refuse?: EngineMessage
  /** The last cell of a table that inserted lines would push off the sheet (its last lines may hold no cell). */
  offSheet?: string
  /** References to deleted tables and columns made #REF!, before the cells move. */
  before(book: BookInput): BookInput
  /** The tables' new ranges and columns, the cells of inserted rows and columns, after the cells moved. */
  after(book: BookInput): BookInput
  /**
   * A reference into a table's data that rows inserted at its edge grow, as Excel inserts rows into a table: from its
   * first data row it stays there, ending on its last data row it takes in rows inserted at the totals row.
   */
  edge?(area: Area): [number, number] | undefined
}

/** Lines inserted (`n` > 0) or deleted (`n` < 0) at index `at` of `axis` on `sheet`. */
export function tableShift(book: BookInput, sheet: string, axis: Axis, at: number, n: number): TableShift {
  const s = sheetIndex(book, sheet)
  const tables = tablesOf(book).list.filter((t) => t.s === s)
  const none: TableShift = { before: (b) => b, after: (b) => b }
  if (!tables.length) return none
  const count = Math.abs(n)
  const last = at + count - 1
  const lines = LINES[axis]
  if (n > 0) {
    const inside = tables.filter((t) => lo(t, axis) <= at && at <= hi(t, axis))
    if (inside.length > 1) return { ...none, refuse: { key: 'tableInsertTwo', lines, tables: inside.map((t) => t.name) } }
  } else {
    const met = tables.filter((t) => lo(t, axis) <= last && hi(t, axis) >= at)
    const cut = met.filter((t) => !(at <= lo(t, axis) && last >= hi(t, axis)))
    // lines that meet two tables go only with both whole (the sets ops and delete-held)
    if (cut.length && met.length > 1) return { ...none, refuse: { key: 'tableDeleteTwo', lines, tables: met.map((t) => t.name) } }
    const head = cut.find((t) => axis === 'r' && t.headers && t.r1 >= at && t.r1 <= last)
    if (head) return { ...none, refuse: { key: 'tableHeader', table: head.name } }
  }
  // what each table becomes: its new range along the axis and its columns, or null when it is deleted
  const plan = new Map<string, { lo: number; hi: number; columns: TableInput['columns']; totals: boolean; grown: number[] } | null>()
  for (const t of tables) {
    const [a, b] = [lo(t, axis), hi(t, axis)]
    const cols = t.input.columns
    if (n > 0) {
      if (at <= a) plan.set(t.name, { lo: a + count, hi: b + count, columns: cols, totals: t.totals, grown: [] })
      else if (at <= b) {
        const grown = Array.from({ length: count }, (_, i) => at + i)
        const columns = axis === 'c' ? [...cols.slice(0, at - t.c1), ...grown.map(() => ({ name: '' })), ...cols.slice(at - t.c1)] : cols
        plan.set(t.name, { lo: a, hi: b + count, columns, totals: t.totals, grown })
      }
      continue
    }
    if (b < at) continue
    if (a > last) {
      plan.set(t.name, { lo: a - count, hi: b - count, columns: cols, totals: t.totals, grown: [] })
      continue
    }
    if (at <= a && last >= b) {
      plan.set(t.name, null)
      continue
    }
    const kept = (i: number) => i < at || i > last
    if (axis === 'c') {
      const columns = cols.filter((_, i) => kept(t.c1 + i))
      const nlo = a < at ? a : at
      plan.set(t.name, { lo: nlo, hi: nlo + columns.length - 1, columns, totals: t.totals, grown: [] })
      continue
    }
    const { first, last: lastData } = rowsOf(t)
    let data = 0
    for (let r = first; r <= lastData; r++) if (kept(r)) data++
    const totals = t.totals && kept(t.r2)
    const nlo = a < at ? a : at
    plan.set(t.name, { lo: nlo, hi: nlo + (t.headers ? 1 : 0) + Math.max(1, data) + (totals ? 1 : 0) - 1, columns: cols, totals, grown: [] })
  }
  if (!plan.size) return none
  const off = n > 0 && tables.find((t) => (plan.get(t.name)?.hi ?? 0) >= LIMIT[axis])
  if (off) return { ...none, offSheet: formatAddress(off.r2, off.c2) }
  const growing = n > 0 && axis === 'r' ? tables.filter((t) => plan.get(t.name)?.grown.length) : []
  const edge = (area: Area): [number, number] | undefined => {
    const x = rectOf(area)
    for (const t of growing) {
      const { first, last } = rowsOf(t)
      if (area.kind === 'cell' || x.c1 < t.c1 || x.c2 > t.c2 || x.r1 < first || x.r2 > last) continue
      if (at === first && x.r1 === first) return [x.r1, x.r2 + count]
      if (at === last + 1 && x.r2 === last) return [x.r1, x.r2 + count]
    }
    return undefined
  }
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
    const names = sheetNames(b)
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
      const first = r.r1 + (t.headers ? 1 : 0)
      let columns = p.columns.map((c) => {
        const next = ti.columns.find((x) => x.name === c.name) ?? c
        // References were shifted with the old first row. Re-anchor the template on a newly inserted first row.
        return next.formula !== undefined && axis === 'r' && n > 0 && t.headers && at === t.r1 + 1 && p.grown.length
          ? { ...next, formula: rewriteFormula(next.formula, sh.name, sh.name, copyTransform(-count, 0), names) }
          : next
      })
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
            cells[addr] = { ...cells[addr], formula: rewriteFormula(c.formula, sh.name, sh.name, copyTransform(row - first, 0), names) }
          })
      const next: TableInput = { ...ti, range: cornersText(r), columns }
      if (p.totals) next.totals = true
      else delete next.totals
      list.push(next)
    }
    const out: SheetInput = withField({ ...sh, cells }, 'tables', list.length ? list : undefined)
    const sheets = [...b.sheets]
    sheets[s] = out
    return { ...b, sheets }
  }
  return { before, after, ...(growing.length ? { edge } : {}) }
}

/**
 * The columns of a table whose every data row the deleted rows `at` … `at + count - 1` take, its header row left: Excel
 * keeps the first data row, cleared, and deletes the rows after it (the set `tables/ops`, del-all-data).
 */
export function emptiedTable(book: BookInput, sheet: string, at: number, count: number): { c1: number; c2: number } | null {
  const s = sheetIndex(book, sheet)
  for (const t of tablesOf(book).list) {
    if (t.s !== s || !t.headers) continue
    const { first, last } = rowsOf(t)
    if (at === first && at + count - 1 >= last) return { c1: t.c1, c2: t.c2 }
  }
  return null
}
