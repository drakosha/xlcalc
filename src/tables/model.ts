/**
 * A book's tables as the engine uses them: each on its sheet's rectangle, its rows split into the header row, the data
 * and the totals row, its columns by name; and the rectangle a structured reference takes (`docs/tables.md`).
 */
import type { BookInput, TableInput } from '../contract.js'
import { parseRange, type Rect } from '../structure/area.js'
import type { TableArea, TableSpec } from '../syntax/table.js'
import { sheetKey } from '../syntax/print.js'

/** Table names, like sheet names, compare without case. */
export const tableKey = sheetKey

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

/** The table holding a cell of sheet `s`. */
export const tableAt = (tables: Tables, s: number, r: number, c: number): Table | undefined => tables.list.find((t) => t.s === s && r >= t.r1 && r <= t.r2 && c >= t.c1 && c <= t.c2)

/** The rows of a table: header row (-1 without one), first and last data rows, totals row (-1 without one). */
export function rowsOf(t: Table): { head: number; first: number; last: number; foot: number } {
  return { head: t.headers ? t.r1 : -1, first: t.r1 + (t.headers ? 1 : 0), last: t.r2 - (t.totals ? 1 : 0), foot: t.totals ? t.r2 : -1 }
}

/**
 * The rectangle a structured reference takes in the formula's row `row`: #REF! for a header or totals row the table
 * does not have and for an unknown column, #VALUE! for the formula's row outside the data, as Excel answers.
 */
export function specRect(t: Table, spec: TableSpec, row: number): Rect | 'REF' | 'VALUE' {
  const { head, first, last, foot } = rowsOf(t)
  const has = (a: TableArea) => spec.areas.includes(a)
  let r1 = first
  let r2 = last
  if (has('all')) [r1, r2] = [t.r1, t.r2]
  else if (has('row')) {
    if (row < first || row > last) return 'VALUE'
    r1 = r2 = row
  } else if (has('headers') && !has('data')) {
    if (head < 0) return 'REF'
    r1 = r2 = head
  } else if (has('totals') && !has('data')) {
    if (foot < 0) return 'REF'
    r1 = r2 = foot
  } else {
    if (has('headers') && head >= 0) r1 = head
    if (has('totals') && foot >= 0) r2 = foot
  }
  if (!spec.cols) return { r1, r2, c1: t.c1, c2: t.c2 }
  const at = spec.cols.map((c) => t.index.get(tableKey(c)))
  if (at.some((i) => i === undefined)) return 'REF'
  const [a, b = a] = at as number[]
  return { r1, r2, c1: t.c1 + Math.min(a, b), c2: t.c1 + Math.max(a, b) }
}
