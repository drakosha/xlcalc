/** The types of the `./xlsx` subpath: reading an xlsx file into a book and writing a book with its values. */
import type { Book, BookInput, CellResult, ComputeResult } from '../contract.js'
import type { Locale } from '../locale/types.js'

/**
 * What the book cannot hold of a file, or what writing changed. The codes:
 *
 * - `richText`: text with formatted runs, read as its plain text.
 * - `name`: a name the book does not take (a built-in `_xlnm.` name, one on a sheet left out), a hidden name read as
 *   a visible one, or a name the writer left out.
 * - `table`: a table the model cannot hold (over another table or a merged area, its name taken), left out; a header
 *   cell that does not hold its column's name as text, given it; a structured reference without a table name.
 * - `externalLink`: a link that is no workbook's (DDE, OLE) or whose part is missing, kept as a book without sheets; a
 *   formula that refers to another workbook the links do not hold (`[3]Sheet!A1` in a file of two), or that the engine
 *   does not parse.
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
 * - `noCachedValues`: written: formulas without their values, which Excel computes when it opens the file: no result was
 *   given (`values: 'omit'` by default), or the result has no value for them (`values: 'preserve'`).
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
  | 'noCachedValues'
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

/**
 * The values a file stores for formulas and for the cells their arrays spilled into (`XlsxRead.cached`): what Excel
 * computed when it saved the file, by sheet in the book's order, in the shape of a result. `cached.sheets[0].cells.B1.value`
 * reads as `compute(book).sheets[0].cells.B1.value` does, so the two compare cell by cell, and `{ book, result: cached }`
 * writes the file again with Excel's values (`writeXlsx`). `compute(read.book, { cached: read.cached })` (and
 * `createSession`) open the book as Excel opens the file: the formulas a circular reference stops hold these values
 * (`CellInput.held`).
 */
export interface XlsxCached {
  sheets: { name: string; cells: Record<string, XlsxCachedCell> }[]
}

/**
 * A stored value: a formula's, with `spill`, the range its array fills (an array formula of the old kind: its area), or
 * a cell of an array's range, with `spillOf`, the address of its formula; as in `CellResult`.
 */
export type XlsxCachedCell = Pick<CellResult, 'value' | 'spill' | 'spillOf'>

/** A file read into a book. */
export interface XlsxRead {
  /** The book, as the engine takes it. */
  book: Book
  /** The values the file stores for formulas and the cells their arrays spilled into, to compare with the engine's result. */
  cached: XlsxCached
  issues: XlsxIssue[]
}

/**
 * What `writeXlsx` writes: a book and the values of its formulas, a session's `getSnapshot()`, `{ book, result:
 * compute(book) }`, or a file read with the values it stored, `{ book: read.book, result: read.cached }`.
 */
export interface XlsxSnapshot {
  book: BookInput
  result?: ComputeResult | XlsxCached
}

/** The options of `readXlsx`. */
export interface ReadXlsxOptions {
  /**
   * The language of the issues' messages: a loaded locale's tag (`'ru-RU'` once `registerLocale(ruRU)` ran) or a locale
   * object (`ruRU`), en-US by default. Codes and the other fields of an issue are the same in every language. A file
   * does not say what language its reader reads, so the host passes the locale the person chose.
   */
  locale?: string | Locale
  /**
   * How large the file's parts may be once inflated: one part, 256 MiB by default, and the parts read together, 1 GiB by
   * default. A file past them is refused, and so is a part that inflates to more than the file says it holds. A host
   * that reads files people upload sets them lower.
   */
  limits?: XlsxLimits
}

/** The bounds of `readXlsx` on a file's parts once inflated (`ReadXlsxOptions.limits`): sizes in bytes, whole numbers of at least 1, and a ratio. */
export interface XlsxLimits {
  /**
   * One part: 536,870,912 (512 MiB) by default. A part is read as one JavaScript string, so a part over about 512 MiB is
   * not read whatever the bound.
   */
  entryBytes?: number
  /** The parts read, together: 1,073,741,824 (1 GiB) by default. */
  totalBytes?: number
  /**
   * How many times its size in the file a part over 16 MiB may inflate to, at least 1: 200 by default. The parts of
   * xlsx files inflate some 5 to 40 times; a zip bomb's, a thousand times.
   */
  ratio?: number
}

/** Options of `writeXlsx`: the language of its issues, which values of formulas the file stores, formulas or values alone. */
export interface WriteXlsxOptions {
  /** The language of the issues' messages, as `readXlsx` takes it (`ReadXlsxOptions.locale`). */
  locale?: string | Locale
  /**
   * The values of formulas the file stores. `require`: each formula's and each value of its array, from the result (an
   * `InputError` when one is missing: the result is not the book's). `omit`: none, and Excel computes them when it opens
   * the file. `preserve`: those the result has; a formula without one is reported as `noCachedValues`. By default
   * `require` with a result and `omit` without one, reported as `noCachedValues` when the book has formulas.
   */
  values?: 'require' | 'omit' | 'preserve'
  /**
   * `formulas` (the default): formulas, with their values by `values`. `values`: values only, a formula's cell gets its
   * value from the result (required then).
   */
  mode?: 'formulas' | 'values'
}

/** What `writeXlsx` answers: the file and what the book held that the file could not. */
export interface XlsxWritten {
  /** The xlsx file. */
  data: Uint8Array
  issues: XlsxIssue[]
}
