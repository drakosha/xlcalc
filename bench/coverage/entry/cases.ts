/**
 * Text typed into cells, for Excel (`run.ts entry <set> [ru|<tag>]`, `bench/windows/run.ts entry <set> [ru|<tag>]`): each
 * case types its text into a cell of its own, column A of the sheet `S`, a cell in General or with a format set before.
 * Excel takes the text as typed in its locale (`Formula2Local`), and the cell is read back: what it holds as Excel saves
 * it (the value or formula and the format, read from the saved file), what it shows and its format codes. Formulas typed
 * read the data in column C, never column A.
 */
import type { BookInput, CellInput, CellValue, NumberFormat } from '../../../src/contract.js'

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
  cell: CellInput | null
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
}

export interface EntrySet {
  buildEntryCases(): EntryCase[]
  /** The texts are typed on the keyboard (Excel for Mac only), to check that Excel takes `Formula2Local` as typing. */
  keys?: boolean
}

/** The data the typed formulas read: a number and a date with the short date format. */
export const DATA: Readonly<Record<string, CellInput>> = { C1: { value: 2 }, C2: { value: 46096, format: 14 } }

/** The book of a part: the cases' cells in A1 down with their formats, and the data. */
export function entryBook(cases: EntryCase[]): BookInput {
  const cells: Record<string, CellInput> = { ...DATA }
  cases.forEach((c, i) => {
    if (c.format !== undefined) cells[`A${i + 1}`] = { format: c.format }
  })
  return { sheets: [{ name: 'S', cells }] }
}

/** Cases of one group: ids `<group>-<n>`. */
export function collector() {
  const out: EntryCase[] = []
  const count = new Map<string, number>()
  const add = (group: string, text: string, format?: NumberFormat) => {
    const n = (count.get(group) ?? 0) + 1
    count.set(group, n)
    out.push({ id: `${group}-${n}`, text, ...(format !== undefined ? { format } : {}) })
  }
  return { out, add }
}
