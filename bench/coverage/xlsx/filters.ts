/**
 * Filters in the forms a file holds them, against Excel. Sheet S holds 2^0 … 2^9 in A2:A11 with keys `x` and `y` beside
 * them and 1 … 16 in D13:H13; sheet F holds SUBTOTAL and AGGREGATE over them, so a sum names the cells it counted
 * (`filterBook`, which the round trip's books of hidden lines take too). `writeXlsx` writes each case's book, then the
 * filter is put in by hand: an autofilter with criteria with and without the sheet's filter mark, one without criteria,
 * the mark alone (an advanced filter), a table's filter, and the writer's criteria that every row meets on a table.
 * `npx tsx bench/coverage/run.ts xlsx filters` has Excel compute and save each file into `test/excel/xlsx/filters/`;
 * `test/xlsx-excel.test.ts` reads what Excel saved and checks that the engine computes the values Excel stored.
 */
import JSZip from 'jszip'
import type { BookInput, CellInput, ColumnInput, RowInput, SheetInput } from '../../../src/contract.js'
import { writeXlsx } from '../../../src/xlsx.js'
import { KEEP_ALL } from '../../../src/xlsx/write/sheet.js'
import { zip } from '../../../src/xlsx/zip/write.js'

/** The forms of a filter: criteria keeping the rows keyed `x`, none, or the writer's criteria that every row meets. */
type Criteria = 'x' | 'none' | 'all'

/** What sheet S has beside its cells. */
export interface FilterSheet {
  rows?: Record<string, RowInput>
  columns?: Record<string, ColumnInput>
  filtered?: true
  /** S!A1:B11 is the table `Data`. */
  table?: true
}

export interface FilterFile {
  id: string
  about: string
  rows: Record<string, RowInput>
  /** An autofilter of the sheet over A1:B11. */
  filter?: Criteria
  /** S!A1:B11 is the table `Data`, with an autofilter of these criteria. */
  table?: Criteria
  /** The sheet's filter mark, `<sheetPr filterMode="1">`. */
  mark?: true
  /** The defined name `_xlnm._FilterDatabase` over A1:B11, as Excel writes beside an advanced filter. */
  database?: true
}

/** The keys of rows 2 to 11: rows 4, 7 and 8 hold `y`. */
const KEYS = 'xxyxxyyxxx'
export const Y_ROWS: Record<string, RowInput> = { 4: { hidden: true }, '7:8': { hidden: true } }

/** The formulas of sheet F, in A1 down. */
export const FILTER_FORMULAS = [
  'SUBTOTAL(9,S!A2:A11)',
  'SUBTOTAL(109,S!A2:A11)',
  'SUBTOTAL(3,S!A2:A11)',
  'SUBTOTAL(103,S!A2:A11)',
  'AGGREGATE(9,5,S!A2:A11)',
  'AGGREGATE(9,4,S!A2:A11)',
  'SUBTOTAL(9,S!D13:H13)',
  'SUBTOTAL(109,S!D13:H13)',
]

export const FILTER_FILES: FilterFile[] = [
  { id: 'filter', about: 'an autofilter keeping x hides rows 4, 7 and 8; the sheet marked filtered', rows: Y_ROWS, filter: 'x', mark: true },
  { id: 'filter-unmarked', about: 'the autofilter keeping x without the sheet’s filter mark', rows: Y_ROWS, filter: 'x' },
  { id: 'filter-none', about: 'an autofilter without criteria and the filter mark, rows 4, 7 and 8 hidden by hand', rows: Y_ROWS, filter: 'none', mark: true },
  { id: 'advanced', about: 'the filter mark without an autofilter and the filter database name, as an advanced filter leaves them, rows 4, 7 and 8 hidden', rows: Y_ROWS, mark: true, database: true },
  { id: 'table', about: 'a table whose filter keeps x hides rows 4, 7 and 8; the sheet marked filtered', rows: Y_ROWS, table: 'x', mark: true },
  { id: 'table-unmarked', about: 'the table’s filter keeping x without the sheet’s filter mark', rows: Y_ROWS, table: 'x' },
  { id: 'table-none', about: 'a table without criteria, rows 4, 7 and 8 hidden by hand', rows: Y_ROWS, table: 'none' },
  { id: 'keep-all-table', about: "the writer's criteria on a table's filter, rows 4, 7 and 8 hidden", rows: Y_ROWS, table: 'all', mark: true },
]

/** The book: sheet S with its data and what `x` gives it, sheet F with the formulas. */
export function filterBook(x: FilterSheet): BookInput {
  const cells: Record<string, CellInput> = { A1: { value: 'v' }, B1: { value: 'k' } }
  for (let i = 0; i < 10; i++) Object.assign(cells, { [`A${i + 2}`]: { value: 2 ** i }, [`B${i + 2}`]: { value: KEYS[i] } })
  ;['D', 'E', 'F', 'G', 'H'].forEach((col, i) => (cells[`${col}13`] = { value: 2 ** i }))
  const s: SheetInput = { name: 'S', cells, ...(x.rows ? { rows: x.rows } : {}), ...(x.columns ? { columns: x.columns } : {}), ...(x.filtered ? { filtered: true } : {}) }
  if (x.table) s.tables = [{ name: 'Data', range: 'A1:B11', columns: [{ name: 'v' }, { name: 'k' }] }]
  return { sheets: [s, { name: 'F', cells: Object.fromEntries(FILTER_FORMULAS.map((f, i) => [`A${i + 1}`, { formula: f }])) }] }
}

const COLUMN: Record<Exclude<Criteria, 'none'>, string> = { x: '<filterColumn colId="1"><filters><filter val="x"/></filters></filterColumn>', all: KEEP_ALL }
const filterXml = (c: Criteria) => `<autoFilter ref="A1:B11">${c === 'none' ? '' : COLUMN[c]}</autoFilter>`

/** The case's file: the book as `writeXlsx` writes it, with the filter put into its parts. */
export async function filterFile(c: FilterFile): Promise<Uint8Array> {
  const parts = await JSZip.loadAsync((await writeXlsx(filterBook({ rows: c.rows, ...(c.table ? { table: true } : {}) }))).data)
  const files = new Map<string, string>()
  for (const [name, f] of Object.entries(parts.files)) if (!f.dir) files.set(name, await f.async('string'))
  const edit = (name: string, fn: (xml: string) => string) => files.set(name, fn(files.get(name)!))
  const sheet = 'xl/worksheets/sheet1.xml'
  if (c.filter) edit(sheet, (x) => x.replace('</sheetData>', `</sheetData>${filterXml(c.filter!)}`))
  if (c.mark) edit(sheet, (x) => x.replace(/(<worksheet[^>]*>)/, '$1<sheetPr filterMode="1"/>'))
  if (c.table) edit('xl/tables/table1.xml', (x) => x.replace(/<autoFilter ref="A1:B11"\/>/, filterXml(c.table!)))
  if (c.database) edit('xl/workbook.xml', (x) => x.replace('</sheets>', '</sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\'S\'!$A$1:$B$11</definedName></definedNames>'))
  const enc = new TextEncoder()
  return zip([...files].map(([name, text]) => ({ name, data: enc.encode(text) })))
}
