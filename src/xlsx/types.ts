/** The types of the `./xlsx` subpath: reading an xlsx file into a book and writing a book with its values. */
import type { BookInput, CellValue, ComputeResult } from '../contract.js'

/**
 * What the book cannot hold of a file, or what writing changed. The codes:
 *
 * - `richText`: text with formatted runs, read as its plain text.
 * - `name`: a name the book does not take (a built-in `_xlnm.` name, one on a sheet left out), a hidden name read as
 *   a visible one, or a name the writer left out.
 * - `table`: a table the model cannot hold (over another table or a merged area, its name taken), left out; a header
 *   cell that does not hold its column's name as text, given it; a structured reference without a table name.
 * - `externalLink`: links to other workbooks, and formulas that reference them (`[1]Sheet!A1`).
 * - `userFunction`: a call of a function the file marks as user-defined (`_xludf.`), a macro or add-in function.
 * - `formula`: a formula the engine does not parse, read as written (its cell computes to an error), or a formula
 *   the writer could not write (its value is written instead).
 * - `errorConstant`: an error typed into a cell as its value, read as a formula of that error.
 * - `errorValue`: an error the model has no code for (`#GETTING_DATA`, `#FIELD!`, `#BLOCKED!` and the like), read
 *   as `#ERROR!`; or written: `#CYCLE!` as 0, as Excel stores a circular reference.
 * - `dataTable`: a what-if data table (`TABLE()`), read as its stored values.
 * - `calculation`: calculation settings the book does not keep: the manual mode, automatic but for data tables,
 *   precision as displayed; an iteration count or change out of range, read as Excel's default.
 * - `style`: what a style cannot hold (pattern and gradient fills, diagonal borders, rotation and the like).
 * - `defaultFont`: a default font other than Calibri 11, which the model assumes.
 * - `defaultWidth`: a sheet's default column width other than Excel's standard of 8 digits, which the model assumes
 *   (Excel for Mac starts its own books at 10).
 * - `numberFormat`: a built-in East Asian date format the model does not have (50 to 58, read as its en-US code), or
 *   one this reader does not know (left out).
 * - `hidden`: hidden sheets, read as visible ones.
 * - `sheetKind`: a chart, dialog or macro sheet, left out.
 * - `cellMetadata`: a cell's rich value (`vm`: a data type, a picture in the cell), read as the value the cell stores.
 * - `autoFilter`: an autofilter's range, criteria and sort, a table's criteria, an advanced filter (the sheet's filter
 *   mode is read: `SheetInput.filtered`); written: the filter mode of a sheet whose tables have no header row, left out.
 * - `dataValidation`: a rule the model cannot hold (no range of the sheet, no formula its type needs, an unknown type),
 *   left out; cells two rules share, kept by the earlier; an IME mode; a list whose items hold commas, read split at
 *   them; written: a rule whose formula the file cannot hold, left out.
 * - `outline`: automatic styles of an outline.
 * - `conditionalFormat`, `comment`, `hyperlink`, `drawing` (pictures, charts, shapes),
 *   `pivotTable`, `protection`, `frozenPanes`, `print`
 *   (headers and footers, page breaks, a page setup made by hand), `macros`, `sparkline`, `connection` (data
 *   connections and queries): parts of the file the model does not have, left
 *   out. View settings (zoom, gridlines, the selection) and a page setup Excel writes by default are not reported.
 * - `locale`: a locale the package does not have, left out of the book.
 * - `meta`: written: cell `meta` has no place in a file.
 * - `textLength`: written: text over Excel's 32,767 characters, cut, and a rule's texts over Excel's limits.
 * - `part`: a part of the file that is missing or cannot be read.
 */
export type XlsxIssueCode =
  | 'richText'
  | 'name'
  | 'table'
  | 'externalLink'
  | 'userFunction'
  | 'formula'
  | 'errorConstant'
  | 'errorValue'
  | 'dataTable'
  | 'calculation'
  | 'style'
  | 'defaultFont'
  | 'defaultWidth'
  | 'numberFormat'
  | 'hidden'
  | 'sheetKind'
  | 'cellMetadata'
  | 'conditionalFormat'
  | 'dataValidation'
  | 'comment'
  | 'hyperlink'
  | 'drawing'
  | 'autoFilter'
  | 'pivotTable'
  | 'protection'
  | 'frozenPanes'
  | 'outline'
  | 'print'
  | 'macros'
  | 'sparkline'
  | 'connection'
  | 'locale'
  | 'meta'
  | 'textLength'
  | 'part'

/** One thing lost or changed, and where: the sheet, and the cells (addresses or ranges) or the name it is about. */
export interface XlsxIssue {
  code: XlsxIssueCode
  /** What was lost or changed, and why. */
  message: string
  sheet?: string
  cells?: string[]
  name?: string
}

/** A file read into a book. */
export interface XlsxRead {
  /** The book, as the engine takes it. */
  book: BookInput
  /**
   * The values the file stores for formulas and for the cells their arrays spilled into, by sheet in the book's
   * order: what Excel computed when it saved the file, to compare with the engine's result.
   */
  cached: { name: string; cells: Record<string, CellValue> }[]
  issues: XlsxIssue[]
}

export interface WriteXlsxOptions {
  /** The book's computed result (`compute`, `Session.result()`): the values of formulas and their spilled arrays. */
  result?: ComputeResult
  /**
   * `formulas` (the default): formulas with their values from `result` (without it, the file asks Excel to compute
   * them on opening). `values`: values only, a formula's cell gets its value from `result` (required then).
   */
  mode?: 'formulas' | 'values'
}

export interface XlsxWritten {
  /** The xlsx file. */
  data: Uint8Array
  issues: XlsxIssue[]
}
