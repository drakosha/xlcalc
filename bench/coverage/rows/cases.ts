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
import { groupCollector } from '../collect.js'

/**
 * How `typed` is typed: Excel's `Value` with the text, `Formula2Local` (Excel 365's typing in the user's language), or
 * on the keyboard (Excel for Mac only): `keys` typed into the cell and Return, `fill` typed into the selection from the
 * cell to `fillTo` and Control+Return, `paste` put on the clipboard and pasted at the cell (Excel for Windows through
 * `Worksheet.Paste`).
 */
export type RowVia = 'value' | 'local' | 'keys' | 'fill' | 'paste'

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
  via?: RowVia
  /** The last column of the selection a `fill` enters its text into, from the typed cell; of a `paste`, the last column it may fill, read after. */
  fillTo?: number
  /** The cells of the row below are read too (a paste may fill it). */
  below?: true
  /** AutoFit after the texts are typed (after `patched` too). */
  fitAfter?: boolean
  /** A font size or a rotation set through the object model on the case's columns `cols` whole, after the texts; the row is fitted only with `fitAfter`. */
  patched?: { cols: [number, number]; size?: number; rotation?: number | 'vertical' }
  /** Rows of the sheet, by number, whose heights are read after too (`RowAnswer.more`). */
  more?: number[]
}

/** What Excel made of a case's row. */
export interface RowAnswer {
  /** The row's height as the object model reports it (`RowHeight`) when the book opened and after, and whether it is the sheet's standard height after. */
  before: number
  height: number
  standard: boolean
  /** The row as Excel saved it: its `<row>` height and flags; null when the file has no height for it. */
  saved: { ht: number; customHeight?: true; hidden?: true } | null
  /** Each of the case's cells after, by column (`<col>+1` for the row below): its content (`Formula`) and whether it wraps its text (`WrapText`). */
  cells: Record<string, { formula: string; wrap: boolean }>
  /** A typed text Excel did not take, by its index: Excel's error. */
  refused?: Record<number, string>
  /** The heights of the rows `RowCase.more` names after, by number. */
  more?: Record<string, number>
  /** Of a case that patches columns, the sheet's `<sheetFormatPr>` as Excel saved it. */
  format?: Record<string, string>
}

export interface RowSet {
  /** The set's cases in a locale (typed texts are written as a user of the locale types them). */
  buildRowCases(tag: string): RowCase[]
}

/** The row of the k-th case of a book (zero-based), 1-based. */
export const caseRow = (k: number) => 2 + 3 * k

/** The first column of the k-th case of a book, zero-based: B, F, J… */
export const caseColumn = (k: number) => 1 + 4 * k

/** Whether a case's row is fitted after its texts: a case that types nothing or patches nothing, else with `fitAfter`. */
export const fitsRow = (c: RowCase): boolean => (!c.typed && !c.patched) || !!c.fitAfter

/** A case whose column has a style or that patches whole columns has a book of its own, as the style reaches every row. */
export const ownBook = (c: RowCase): boolean => !!c.patched || Object.values(c.columns ?? {}).some((x) => x.look)

/** The columns a case has a cell in, its cells' and those typed into (a fill's selection whole). */
export function caseCols(c: RowCase): number[] {
  const typed = (c.typed ?? []).flatMap(([col]) => (c.fillTo === undefined ? [col] : Array.from({ length: c.fillTo - col + 1 }, (_, k) => col + k)))
  return [...new Set([...Object.keys(c.cells ?? {}).map(Number), ...typed])].sort((a, b) => a - b)
}

/** The cells a case reads after, as `[row offset, column]`: its columns, and with `below` those of the row below. */
export const caseCells = (c: RowCase): [number, number][] => [0, ...(c.below ? [1] : [])].flatMap((dr) => caseCols(c).map((col): [number, number] => [dr, col]))

/** The key of a read cell in `RowAnswer.cells`. */
export const cellKey = ([dr, col]: [number, number]) => (dr ? `${col}+${dr}` : String(col))

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
export const collector = () => groupCollector<RowCase>()

/** A wrapping cell's style, a font size over a style, a row's first cell alone with its look. */
export const WRAP: CellStyle = { alignment: { wrap: true } }
export const size = (n: number, more: CellStyle = {}): CellStyle => ({ ...more, font: { size: n } })
export const one = (value: RowCell['value'], look?: CellStyle): Record<number, RowCell> => ({ 0: { value, ...(look ? { look } : {}) } })
