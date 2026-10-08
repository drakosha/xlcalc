/**
 * Table cases for Excel (`run.ts tables <set>`): a small book with tables, steps taken after Excel computes it (cells
 * typed, structure and table operations), and the rectangles read at the end, with the tables of every sheet.
 */
import type { Val } from '../formulajs.js'
import type { BookInput, CellInput, RangeRef, StructureOp, TableColumn, TableInput } from '../../../src/contract.js'

/** A step after the computation, Excel recalculating after each: a cell typed or cleared, or an operation. */
export type TableStep =
  /**
   * A number or a text typed, a formula entered (`{ formula }`, as Excel 365 enters it; `{ local }` typed in the user's
   * language, in a set run in ru-RU), null to clear the cell.
   */
  | { edit: { sheet: string; cell: string; value: number | string | null | { formula: string } | { local: string } } }
  | { op: StructureOp }

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
  book: BookInput
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
  /** Excel did not load the book (the repair window, declined). */
  rejected?: true
}

/** A cell: `=…` a formula, else a value. */
export type CellSpec = number | string | boolean | CellInput

export const cell = (v: CellSpec): CellInput => (typeof v === 'string' && v.startsWith('=') ? { formula: v.slice(1) } : typeof v === 'object' ? v : { value: v })

export const cellsOf = (cells: Record<string, CellSpec>): Record<string, CellInput> => Object.fromEntries(Object.entries(cells).map(([a, v]) => [a, cell(v)]))

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
): { cells: Record<string, CellInput>; table: TableInput } {
  const m = /^([A-Z]+)(\d+)$/.exec(at)!
  const col0 = m[1].split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1
  const row0 = Number(m[2])
  const letter = (i: number) => {
    let n = col0 + i + 1
    let s = ''
    while (n > 0) {
      s = String.fromCharCode(65 + ((n - 1) % 26)) + s
      n = Math.floor((n - 1) / 26)
    }
    return s
  }
  const cells: Record<string, CellInput> = {}
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
export function sales(): { cells: Record<string, CellInput>; table: TableInput } {
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
export function book(...sheets: [string, Record<string, CellSpec>, ...{ cells: Record<string, CellInput>; table: TableInput }[]][]): BookInput {
  return {
    sheets: sheets.map(([name, cells, ...tables]) => ({
      name,
      cells: { ...Object.assign({}, ...tables.map((t) => t.cells)), ...cellsOf(cells) },
      ...(tables.length ? { tables: tables.map((t) => t.table) } : {}),
    })),
  }
}

/** Answers for predictions. */
export const n = (v: number): Val => ({ t: 'n', v })
export const e = (v: Extract<Val, { t: 'e' }>['v']): Val => ({ t: 'e', v })
export const s = (v: string): Val => ({ t: 's', v })

/** Rectangles of one sheet. */
export const on = (sheet: string, ...ranges: string[]): RangeRef[] => ranges.map((range) => ({ sheet, range }))
