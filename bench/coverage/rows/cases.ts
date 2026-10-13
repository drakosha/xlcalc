/**
 * Row heights for Excel (`run.ts rows <set> [ru|<tag>]`, `bench/windows/run.ts rows <set> [ru|<tag>]`): each case is a
 * row of the sheet `S` with its cells, and Excel either fits the row (AutoFit, a double click on the border below its
 * number) or types texts into its cells, fitting the row after when the case says so; the row's height is read before
 * and after, with the wrapping of its cells, and the book is saved so the height Excel stores is read from the file.
 * Case k takes row 2 + 3k and columns 4k + 1 to 4k + 4 of its own (B:E, F:I…), so its columns' widths are its alone, a
 * cell can be merged with the empty cell to its right or below, and no other case's cell stands in its row.
 */
import type { BookInput, CellInput, CellStyle, ColumnInput, RowInput } from '../../../src/contract.js'
import { formatAddress, indexToCol } from '../../../src/model.js'
import { internStyle } from '../../../src/style-table.js'
import { takenPart, type StoredCell } from '../stored.js'
import type { Via } from '../fit/cases.js'

/** A cell of a case: a stored cell with its style by its fields, which the book's table takes. */
export type RowCell = Omit<StoredCell, 'style'> & { look?: CellStyle }

export interface RowCase {
  id: string
  /** The row's cells by the case's column, 0 to 2. */
  cells?: Record<number, RowCell>
  /** The case's columns by their place: a width, hidden, a style by its fields. */
  columns?: Record<number, Omit<ColumnInput, 'style'> & { look?: CellStyle }>
  /** The row before: a height set by hand (xlsx `customHeight`), hidden, a style by its fields. */
  row?: Omit<RowInput, 'style'> & { look?: CellStyle }
  /** Cells merged with the cell to their right, by column. */
  merged?: number[]
  /** Cells merged with the cell below them, by column. */
  mergedDown?: number[]
  /** Texts typed into the row's cells in turn, each by column; without them the row is fitted (AutoFit). */
  typed?: [col: number, text: string][]
  /** How `typed` is typed; `local` without it. */
  via?: Via
  /** AutoFit after the texts are typed. */
  fitAfter?: boolean
}

/** What Excel made of a case's row. */
export interface RowAnswer {
  /** The row's height as the object model reports it (`RowHeight`) when the book opened and after, and whether it is the sheet's standard height after. */
  before: number
  height: number
  standard: boolean
  /** The row as Excel saved it: its `<row>` height and flags; null when the file has no height for it. */
  saved: { ht: number; customHeight?: true; hidden?: true } | null
  /** Whether each of the case's cells wraps its text after (`WrapText`), by column. */
  wrap: Record<number, boolean>
  /** A typed text Excel did not take, by its index: Excel's error. */
  refused?: Record<number, string>
}

export interface RowSet {
  /** The set's cases in a locale (typed texts are written as a user of the locale types them). */
  buildRowCases(tag: string): RowCase[]
}

/** The row of the k-th case of a book (zero-based), 1-based. */
export const caseRow = (k: number) => 2 + 3 * k

/** The first column of the k-th case of a book, zero-based: B, F, J… */
export const caseColumn = (k: number) => 1 + 4 * k

/** The columns a case has a cell in, its cells' and those typed into. */
export const caseCols = (c: RowCase): number[] => [...new Set([...Object.keys(c.cells ?? {}).map(Number), ...(c.typed ?? []).map(([col]) => col)])].sort((a, b) => a - b)

/** The book of a part, as the engine takes it: each case's cells, columns, row and merged areas, styles interned into one table. */
export function rowsBook(cases: RowCase[]): BookInput {
  let styles: CellStyle[] = []
  const styleOf = (look: CellStyle) => {
    const r = internStyle(styles, look)
    styles = r.styles
    return r.index
  }
  const cells: Record<string, CellInput> = {}
  const columns: Record<string, ColumnInput> = {}
  const rows: Record<string, RowInput> = {}
  const merges: string[] = []
  cases.forEach((c, k) => {
    const r = caseRow(k) - 1
    const c0 = caseColumn(k)
    for (const [col, { look, ...cell }] of Object.entries(c.cells ?? {})) cells[formatAddress(r, c0 + Number(col))] = { ...takenPart(cell), ...(look ? { style: styleOf(look) } : {}) }
    for (const [col, { look, ...x }] of Object.entries(c.columns ?? {})) columns[indexToCol(c0 + Number(col))] = { ...x, ...(look ? { style: styleOf(look) } : {}) }
    if (c.row) {
      const { look, ...row } = c.row
      rows[String(r + 1)] = { ...row, ...(look ? { style: styleOf(look) } : {}) }
    }
    for (const col of c.merged ?? []) merges.push(`${formatAddress(r, c0 + col)}:${formatAddress(r, c0 + col + 1)}`)
    for (const col of c.mergedDown ?? []) merges.push(`${formatAddress(r, c0 + col)}:${formatAddress(r + 1, c0 + col)}`)
  })
  return { sheets: [{ name: 'S', cells, columns, rows, ...(merges.length ? { merges } : {}) }], styles }
}

/** Cases of one group: ids `<group>-<n>`. */
export function collector() {
  const out: RowCase[] = []
  const count = new Map<string, number>()
  const add = (group: string, c: Omit<RowCase, 'id'>) => {
    const n = (count.get(group) ?? 0) + 1
    count.set(group, n)
    out.push({ id: `${group}-${n}`, ...c })
  }
  return { out, add }
}
