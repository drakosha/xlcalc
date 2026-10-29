/**
 * Text typed into cells, for Excel (`run.ts entry <set> [ru|<tag>]`, `bench/windows/run.ts entry <set> [ru|<tag>]`): each
 * case types its text into a cell of its own, column A of the sheet `S`, a cell in General or with a format set before.
 * Excel takes the text as typed in its locale (`Formula2Local`), and the cell is read back: what it holds as Excel saves
 * it (the value or formula and the format, read from the saved file), what it shows and its format codes. Formulas typed
 * read the data in column C, never column A.
 */
import type { BookInput, CellInput, CellValue, ColumnInput, NameInput, NumberFormat, SheetInput } from '../../../src/contract.js'
import type { StoredCell } from '../stored.js'
import { groupCollector } from '../collect.js'

export interface EntryCase {
  id: string
  /** The text typed, as a user types it into the cell. */
  text: string
  /** The cell's format before the text is typed (an en-US code, or a built-in format by its number); General without it. */
  format?: NumberFormat
}

/** What Excel made of a case's text. */
export interface EntryAnswer {
  /** Excel did not take the text (a formula it refuses): its error, the cell left as it was. */
  refused?: string
  /** The cell as Excel saved it, read back by `readXlsx`: its value or formula and its format; null when the cell is blank. */
  cell: StoredCell | null
  /** A formula's value as Excel saved it. */
  computed?: CellValue
  /** What the cell shows (`Range.Text`) in the default column width, and its format codes (`NumberFormat`, `NumberFormatLocal`). */
  text: string
  code: string
  local: string
  /** The cell's prefix character (`PrefixCharacter`): `'` after a typed apostrophe; absent when empty. */
  prefix?: string
  /** The number the object model gives as the cell's value (`Value2` on Windows, `value` on the Mac); absent for anything else. */
  value2?: number
  /**
   * For a set of `edit`: the cell's `Formula` and `FormulaLocal` (the number itself for a date, a time or a percent), and
   * on the Mac what its formula bar shows, read through the window's accessibility (`Formula Bar. A,1 <text>`), whose text
   * is `formulaBar`.
   */
  formula?: string
  formulaLocal?: string
  bar?: string
}

export interface EntrySet {
  buildEntryCases(): EntryCase[]
  /** The texts are typed on the keyboard (Excel for Mac only), to check that Excel takes `Formula2Local` as typing. */
  keys?: boolean
  /** Each cell's `Formula` and `FormulaLocal` are read too, for `entryText`. */
  edit?: boolean
  /** The data the formulas read in place of `DATA`: the cells of `S` beside the cases', the sheets after it, the names. */
  data?: EntryData
}

/** A set's data: cells of the sheet `S` outside column A and its columns, more sheets, defined names. */
export interface EntryData {
  cells: Record<string, CellInput>
  columns?: Record<string, ColumnInput>
  sheets?: SheetInput[]
  names?: NameInput[]
}

/** The text of the formula bar from its description (`Formula Bar. A,1 1/31/2025`, in the interface's language). */
export const formulaBar = (bar: string): string => bar.replace(/^[^.]*\. [A-Z]{1,3},\d+ ?/, '')

/** The data the typed formulas read: a number and a date with the short date format. */
export const DATA: Readonly<Record<string, CellInput>> = { C1: { value: 2 }, C2: { value: 46096, numberFormat: 14 } }

/** The book of a part: the cases' cells in A1 down with their formats, and the data (the set's, else `DATA`). */
export function entryBook(cases: EntryCase[], data: EntryData = { cells: DATA }): BookInput {
  const cells: Record<string, CellInput> = { ...data.cells }
  cases.forEach((c, i) => {
    if (c.format !== undefined) cells[`A${i + 1}`] = { numberFormat: c.format }
  })
  return { sheets: [{ name: 'S', cells, ...(data.columns ? { columns: data.columns } : {}) }, ...(data.sheets ?? [])], ...(data.names ? { names: data.names } : {}) }
}

/** Cases of one group: ids `<group>-<n>`. */
/** Cases of texts typed into General cells (by group), then into formatted cells (`[group, format, texts]`). */
export function entryCases(general: Readonly<Record<string, readonly string[]>>, formatted: readonly (readonly [string, NumberFormat, readonly string[]])[] = []): EntryCase[] {
  const { out, add } = collector()
  for (const [group, texts] of Object.entries(general)) for (const t of texts) add(group, t)
  for (const [group, format, texts] of formatted) for (const t of texts) add(group, t, format)
  return out
}

export function collector() {
  const { out, add: put } = groupCollector<EntryCase>()
  const add = (group: string, text: string, format?: NumberFormat) => put(group, { text, ...(format !== undefined ? { format } : {}) })
  return { out, add }
}
