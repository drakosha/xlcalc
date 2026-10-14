/**
 * Table cases for Excel (`run.ts tables <set>`): a small book with tables, steps taken after Excel computes it (cells
 * typed, structure and table operations), and the rectangles read at the end, with the tables of every sheet.
 */
import type { Val } from '../formulajs.js'
import type { BookInput, NameInput, RangeRef, StructureOp, TableColumn, TableInput } from '../../../src/contract.js'
import { colToIndex, indexToCol } from '../../../src/model.js'
import type { StoredBook, StoredCell } from '../stored.js'

/** A step after the computation, Excel recalculating after each: a cell typed or cleared, an operation, a name added. */
export type TableStep =
  /**
   * A number or a text typed, a formula entered (`{ formula }`, as Excel 365 enters it; `{ local }` typed in the user's
   * language, in a set run in ru-RU), null to clear the cell.
   */
  | { edit: { sheet: string; cell: string; value: number | string | null | { formula: string } | { local: string } } }
  | { op: StructureOp }
  /** A defined name added through Excel's object model (`Names.Add`), which refuses a name Excel does not take; the engine's `setNames`. */
  | { name: NameInput }

/** A table as Excel lists it after the steps: its sheet, name, range, rows shown, column names. */
export interface TableShape {
  sheet: string
  name: string
  range: string
  headers: boolean
  totals: boolean
  columns: string[]
}

export interface TableCase {
  id: string
  /** What the case checks. */
  about: string
  /** Formulas in the engine's spelling, written as `writeXlsx` writes them. */
  book: StoredBook
  steps?: TableStep[]
  /** Rectangles read after the steps, where the steps left their cells. */
  read: RangeRef[]
}

export interface TableSet {
  buildTableCases(): TableCase[]
  /** Formulas are read in the user's language too, for a set run in Excel ru-RU. */
  local?: boolean
}

/** The answer Excel gave a case. */
export interface TableAnswer {
  /** Non-blank cells of the read rectangles (`S!A1`). */
  values: Record<string, Val>
  /** Formulas as Excel shows them (`=SUM(Sales[Qty])`), and in the user's language for a local set. */
  formulas: Record<string, string>
  local?: Record<string, string>
  tables: TableShape[]
  /** The steps Excel refused, by index: its error. */
  refused?: Record<number, string>
  /** For a case that adds names: the visible defined names after the steps, a sheet's own as `Sheet!name`. */
  names?: string[]
  /** Excel did not load the book (the repair window, declined). */
  rejected?: true
}

/** A cell: `=…` a formula, else a value. */
export type CellSpec = number | string | boolean | StoredCell

export const cell = (v: CellSpec): StoredCell => (typeof v === 'string' && v.startsWith('=') ? { formula: v.slice(1) } : typeof v === 'object' ? v : { value: v })

export const cellsOf = (cells: Record<string, CellSpec>): Record<string, StoredCell> => Object.fromEntries(Object.entries(cells).map(([a, v]) => [a, cell(v)]))

/**
 * A table from its header cells' row: `table('Sales', 'A1', ['Region', 'Qty'], rows)` puts the names in A1:B1 and the rows
 * below, and gives the table over them; `opts` adds a totals row (its cells), leaves the header row out, or sets columns'
 * formulas and totals.
 */
export function table(
  name: string,
  at: string,
  names: string[],
  rows: CellSpec[][],
  opts: { headers?: false; totals?: CellSpec[]; columns?: Partial<TableColumn>[] } = {},
): { cells: Record<string, StoredCell>; table: TableInput } {
  const m = /^([A-Z]+)(\d+)$/.exec(at)!
  const col0 = colToIndex(m[1])
  const row0 = Number(m[2])
  const letter = (i: number) => indexToCol(col0 + i)
  const cells: Record<string, StoredCell> = {}
  let r = row0
  if (opts.headers !== false) {
    names.forEach((h, i) => (cells[`${letter(i)}${r}`] = { value: h }))
    r++
  }
  for (const row of [...rows, ...(opts.totals ? [opts.totals] : [])]) {
    row.forEach((v, i) => v !== null && v !== '' && (cells[`${letter(i)}${r}`] = cell(v)))
    r++
  }
  const columns = names.map((n, i) => ({ name: n, ...opts.columns?.[i] }))
  const t: TableInput = { name, range: `${at}:${letter(names.length - 1)}${r - 1}`, columns, ...(opts.headers === false ? { headers: false } : {}), ...(opts.totals ? { totals: true } : {}) }
  return { cells, table: t }
}

/** Sales on S!A1:D5: Region, Qty, Price, and Total, a calculated column of Qty times Price. */
export function sales(): { cells: Record<string, StoredCell>; table: TableInput } {
  const calc = '=[@Qty]*[@Price]'
  return table(
    'Sales',
    'A1',
    ['Region', 'Qty', 'Price', 'Total'],
    [
      ['North', 2, 10, calc],
      ['South', 3, 20, calc],
      ['North', 5, 30, calc],
      ['East', 1, 40, calc],
    ],
    { columns: [{}, {}, {}, { formula: 'Sales[[#This Row],[Qty]]*Sales[[#This Row],[Price]]' }] },
  )
}

/** A book of sheets: each a name, its cells (`=…` formulas) and its tables (from `table`). */
export function book(...sheets: [string, Record<string, CellSpec>, ...{ cells: Record<string, StoredCell>; table: TableInput }[]][]): StoredBook {
  return {
    sheets: sheets.map(([name, cells, ...tables]) => ({
      name,
      cells: { ...Object.assign({}, ...tables.map((t) => t.cells)), ...cellsOf(cells) },
      ...(tables.length ? { tables: tables.map((t) => t.table) } : {}),
    })),
  }
}

/** Answers for predictions. */
export { n, e, s } from '../vals.js'

/** Rectangles of one sheet. */
export const on = (sheet: string, ...ranges: string[]): RangeRef[] => ranges.map((range) => ({ sheet, range }))

/** A formula typed into a cell, of sheet S unless another is given. */
export const typedFormula = (cell: string, formula: string, sheet = 'S'): TableStep => ({ edit: { sheet, cell, value: { formula } } })

/** Text typed into a cell in the book's language, of sheet S unless another is given. */
export const typedLocal = (cell: string, text: string, sheet = 'S'): TableStep => ({ edit: { sheet, cell, value: { local: text } } })

/** A defined name added, global unless a sheet scopes it. */
export const addedName = (name: string, ref: string, scope?: string): TableStep => ({ name: { name, ref, ...(scope ? { scope } : {}) } })

/** The book in ru-RU. */
export const inRu = (b: BookInput): BookInput => ({ ...b, locale: 'ru-RU' })
