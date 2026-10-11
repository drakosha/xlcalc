/**
 * Column widths for Excel (`run.ts fit <set> [ru|<tag>]`, `bench/windows/run.ts fit <set> [ru|<tag>]`): each case is a
 * column of the sheet `S` with its cells, and Excel either fits the column (AutoFit, a double click on the border of its
 * header) or types texts into its cells in turn; then the column's width is read, and the book is saved so the width
 * Excel stores is read from the file. Cases take every other column from B, so a cell can be merged with the empty
 * column to its right; row `HIDDEN` is hidden across the sheet, and only the cases about hidden rows use it.
 */
import type { BookInput, CellInput, CellStyle, ColumnInput } from '../../../src/contract.js'
import { formatAddress } from '../../../src/model.js'
import { takenPart, type StoredCell, type StoredColumn } from '../stored.js'

/** How a text is typed: Excel's `Value` with the text, `Formula2Local` (Excel 365's typing in the user's language), or keys (Excel for Mac only). */
export type Via = 'value' | 'local' | 'keys'

export interface FitCase {
  id: string
  /** The column's cells by row number, as a stored case names their fields (errors come from formulas). */
  cells?: Record<number, StoredCell>
  /** The column before: a width set by hand (xlsx `customWidth`), its format, hidden; the default width without them. */
  column?: StoredColumn
  /** Rows whose cell is merged with the cell to its right. */
  merged?: number[]
  /** Rows merged down: the cell of the first with the cells below it to the last. */
  mergedDown?: [number, number]
  /** Texts typed into cells in turn, each by row; without them the column is fitted (AutoFit). */
  typed?: [row: number, text: string][]
  /** How `typed` is typed; `local` without it. */
  via?: Via
  /** AutoFit before the texts are typed. */
  fitFirst?: boolean
}

/** What Excel made of a case's column. */
export interface FitAnswer {
  /** The column's width after, as the object model reports it (`ColumnWidth`), in points (`Width`), and whether it is the sheet's standard width. */
  columnWidth: number
  points: number
  standard: boolean
  /** The column as Excel saved it: its `<col>` width and flags; null when the file has none for it. */
  saved: { width: number; customWidth?: true; bestFit?: true; hidden?: true } | null
  /** What the case's cells show after, by row. */
  texts: Record<number, string>
  /** A typed text Excel did not take, by its index: Excel's error. */
  refused?: Record<number, string>
}

export interface FitSet {
  /** The set's cases in a locale (typed texts are written as a user of the locale types them). */
  buildFitCases(tag: string): FitCase[]
}

/** The row hidden across the sheet. */
export const HIDDEN = 30

/** The styles cases may give cells: bold, a larger size. */
export const STYLES: CellStyle[] = [{ font: { bold: true } }, { font: { size: 16 } }]

/** The column of the k-th case of a book (zero-based): B, D, F… */
export const caseColumn = (k: number) => 1 + 2 * k

/** The rows a case shows something in: its cells and the cells typed into. */
export const caseRows = (c: FitCase): number[] => [...new Set([...Object.keys(c.cells ?? {}).map(Number), ...(c.typed ?? []).map(([r]) => r)])].sort((a, b) => a - b)

/** The book of a part, as the engine takes it: each case's cells, column and merged areas in its column of the sheet `S`. */
export function fitBook(cases: FitCase[]): BookInput {
  const cells: Record<string, CellInput> = {}
  const columns: Record<string, ColumnInput> = {}
  const merges: string[] = []
  cases.forEach((c, k) => {
    const col = caseColumn(k)
    for (const [r, cell] of Object.entries(c.cells ?? {})) cells[formatAddress(Number(r) - 1, col)] = takenPart(cell)
    if (c.column) columns[formatAddress(0, col).replace(/\d+$/, '')] = takenPart(c.column)
    for (const r of c.merged ?? []) merges.push(`${formatAddress(r - 1, col)}:${formatAddress(r - 1, col + 1)}`)
    if (c.mergedDown) merges.push(`${formatAddress(c.mergedDown[0] - 1, col)}:${formatAddress(c.mergedDown[1] - 1, col)}`)
  })
  return { sheets: [{ name: 'S', cells, columns, rows: { [HIDDEN]: { hidden: true } }, ...(merges.length ? { merges } : {}) }], styles: STYLES }
}

/** Cases of one group: ids `<group>-<n>`. */
export function collector() {
  const out: FitCase[] = []
  const count = new Map<string, number>()
  const add = (group: string, c: Omit<FitCase, 'id'>) => {
    const n = (count.get(group) ?? 0) + 1
    count.set(group, n)
    out.push({ id: `${group}-${n}`, ...c })
  }
  return { out, add }
}
