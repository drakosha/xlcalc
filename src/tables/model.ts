/**
 * A book's tables as the engine uses them: each on its sheet's rectangle, its rows split into the header row, the data
 * and the totals row, its columns by name; and the rectangle a structured reference takes (`docs/tables.md`).
 */
import type { BookInput, TableInput } from '../contract.js'
import { hasCell, parseRange, type Rect } from '../structure/area.js'
import type { TableArea, TableSpec } from '../syntax/table.js'
import { sheetKey } from '../syntax/print.js'

/** Table names, like sheet names, compare without case. */
export const tableKey = sheetKey

/** The first name `${base}${n}` (n counted from `from`) that no name of `taken` has, case-insensitively. */
export function numbered(base: string, taken: string[], from: number): string {
  let n = from
  while (taken.some((x) => tableKey(x) === tableKey(`${base}${n}`))) n++
  return `${base}${n}`
}

export interface Table extends Rect {
  name: string
  /** The sheet's index. */
  s: number
  headers: boolean
  totals: boolean
  /** Column names as the table spells them, left to right. */
  cols: string[]
  /** Column index by key (`tableKey`). */
  index: Map<string, number>
  input: TableInput
}

/** The tables of a book, by key and by sheet. */
export interface Tables {
  list: Table[]
  byName: Map<string, Table>
}

/** The tables of a book whose input was checked (`check.ts`). */
export function tablesOf(book: BookInput): Tables {
  const list: Table[] = []
  book.sheets.forEach((sheet, s) => {
    for (const t of sheet.tables ?? []) {
      const x = parseRange(t.range)!
      const cols = t.columns.map((c) => c.name)
      list.push({ ...x, name: t.name, s, headers: t.headers !== false, totals: t.totals === true, cols, index: new Map(cols.map((c, i) => [tableKey(c), i])), input: t })
    }
  })
  return { list, byName: new Map(list.map((t) => [tableKey(t.name), t])) }
}

/** Tables as a formula's check sees them (`validate`, `bindTables`): names and columns, made once for a book's tables. */
const named = new WeakMap<Tables, { name: string; columns: string[] }[]>()
export function tableNames(tables: Tables): { name: string; columns: string[] }[] {
  let out = named.get(tables)
  if (!out) named.set(tables, (out = tables.list.map((t) => ({ name: t.name, columns: t.cols }))))
  return out
}

/** The table holding a cell of sheet `s`. */
export const tableAt = (tables: Tables, s: number, r: number, c: number): Table | undefined => tables.list.find((t) => t.s === s && hasCell(t, r, c))

/** The rows of a table: header row (-1 without one), first and last data rows, totals row (-1 without one). */
export function rowsOf(t: Table): { head: number; first: number; last: number; foot: number } {
  return { head: t.headers ? t.r1 : -1, first: t.r1 + (t.headers ? 1 : 0), last: t.r2 - (t.totals ? 1 : 0), foot: t.totals ? t.r2 : -1 }
}

/**
 * The rows a structured reference's areas take, every area but the formula's row; #REF! for a header or totals row the
 * table does not have. The formula's row (`row`) stays to its caller, which checks it against the data.
 */
export function specRows(t: Table, spec: TableSpec): { r1: number; r2: number } | 'REF' {
  const { head, first, last, foot } = rowsOf(t)
  const has = (a: TableArea) => spec.areas.includes(a)
  if (has('all')) return { r1: t.r1, r2: t.r2 }
  if (has('headers') && !has('data')) return head < 0 ? 'REF' : { r1: head, r2: head }
  if (has('totals') && !has('data')) return foot < 0 ? 'REF' : { r1: foot, r2: foot }
  return { r1: has('headers') && head >= 0 ? head : first, r2: has('totals') && foot >= 0 ? foot : last }
}

/**
 * The rectangle a structured reference takes in the formula's row `row`: #REF! for a header or totals row the table
 * does not have and for an unknown column, #VALUE! for the formula's row outside the data, as Excel answers.
 */
export function specRect(t: Table, spec: TableSpec, row: number): Rect | 'REF' | 'VALUE' {
  let r1: number
  let r2: number
  if (spec.areas.includes('row')) {
    const { first, last } = rowsOf(t)
    if (row < first || row > last) return 'VALUE'
    ;[r1, r2] = [row, row]
  } else {
    const rows = specRows(t, spec)
    if (rows === 'REF') return 'REF'
    ;({ r1, r2 } = rows)
  }
  if (!spec.cols) return { r1, r2, c1: t.c1, c2: t.c2 }
  const at = spec.cols.map((c) => t.index.get(tableKey(c)))
  if (at.some((i) => i === undefined)) return 'REF'
  const [a, b = a] = at as number[]
  return { r1, r2, c1: t.c1 + Math.min(a, b), c2: t.c1 + Math.max(a, b) }
}
