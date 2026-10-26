/** The types of the `./xlsx` subpath: reading an xlsx file into a book and writing a book with its values. */
import type { Book, BookInput, CellResult, ComputeResult } from '../contract.js'
import type { Locale } from '../locale/types.js'

/**
 * What the workbook cannot hold from a file, or what writing changed. The codes:
 *
 * - `richText`: text with formatted runs, read as plain text.
 * - `name`: a name the workbook does not accept (a built-in `_xlnm.` name, one on a sheet that was left out), or a name
 *   the writer left out.
 * - `table`: a table the model cannot hold (overlapping another table or a merged area, or with a name already used),
 *   left out; a header cell that does not hold its column's name as text, given that name; a structured reference
 *   without a table name.
 * - `externalLink`: a link that is not to a workbook (DDE, OLE) or whose part is missing, kept as a workbook without
 *   sheets; a formula that refers to another workbook not among the links (`[3]Sheet!A1` in a file with two), or that
 *   the engine cannot parse.
 * - `userFunction`: a call to a function the file marks as user-defined (`_xludf.`), a macro or add-in function.
 * - `formula`: a formula the engine cannot parse, read as written (its cell calculates to an error), or a formula the
 *   writer could not write (its value is written instead, or an empty cell when it has none).
 * - `errorConstant`: an error typed into a cell as its value, read as a formula of that error.
 * - `errorValue`: an error the model has no code for (`#GETTING_DATA`, `#FIELD!`, `#BLOCKED!` and the like), read
 *   as `#ERROR!`; or, when writing, `#CYCLE!` written as 0, as Excel stores a circular reference, and `#ERROR!`
 *   written as `#NULL!`, which Excel has.
 * - `dataTable`: a what-if data table (`TABLE()`), read as its stored values.
 * - `calculation`: calculation settings the workbook does not keep: manual mode, automatic except for data tables,
 *   precision as displayed; an iteration count or maximum change out of range, read as Excel's default.
 * - `style`: what a style cannot hold (pattern and gradient fills, diagonal borders, rotation and the like).
 * - `defaultFont`: a default font other than Calibri 11, kept in `BookInput.font` and measured as Calibri 11.
 * - `defaultWidth`: a sheet's default column width other than Excel's standard of 8 digits, which the model assumes
 *   (Excel for Mac starts its own workbooks at 10).
 * - `numberFormat`: a built-in East Asian date format the model does not have (50 to 58, read as its en-US code), or
 *   one this reader does not know (left out).
 * - `hidden`: hidden sheets, read as visible ones.
 * - `sheetKind`: a chart, dialog or macro sheet, left out.
 * - `cellMetadata`: a cell's rich value (`vm`: a data type, a picture in the cell), read as the value the cell stores.
 * - `autoFilter`: an AutoFilter's range, criteria and sort, a table's criteria, an advanced filter (the sheet's filter
 *   mode is read: `SheetInput.filtered`); when writing, the filter mode of a sheet whose tables have no header row,
 *   left out.
 * - `dataValidation`: a rule the model cannot hold (no range on the sheet, no formula that its type requires, an
 *   unknown type), left out; cells shared by two rules, kept by the earlier one; an IME mode; a list whose items contain
 *   commas, read as split at them; when writing, a rule whose formula the file cannot hold, left out.
 * - `outline`: automatic styles of an outline.
 * - `conditionalFormat`, `comment`, `hyperlink`, `drawing` (pictures, charts, shapes),
 *   `pivotTable`, `protection`, `frozenPanes`, `print`
 *   (headers and footers, page breaks, a manually made page setup), `macros`, `sparkline`, `connection` (data
 *   connections and queries): parts of the file the model does not have, left
 *   out. View settings (zoom, gridlines, the selection) and a page setup Excel writes by default are not reported.
 * - `locale`: a locale the package does not have, left out of the workbook.
 * - `meta`: when writing, cell `meta` has no place in a file.
 * - `textLength`: when writing, text longer than Excel's 32,767 characters, truncated, and a rule's texts beyond Excel's
 *   limits.
 * - `noCachedValues`: when writing, formulas without their values, which Excel calculates when it opens the file: no
 *   result was given (`values: 'omit'` by default), or the result has no value for them (`values: 'preserve'`).
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

/** One item lost or changed, and where: the sheet, and the cells (addresses or ranges) or the name it concerns. */
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
 * calculated when it saved the file, by sheet in the workbook's order, in the shape of a result.
 * `cached.sheets[0].cells.B1.value` is read the same way as `compute(book).sheets[0].cells.B1.value`, so the two can be
 * compared cell by cell, and `{ book, result: cached }` writes the file again with Excel's values (`writeXlsx`).
 * `compute(read.book, { cached: read.cached })` (and `createSession`) opens the workbook as Excel opens the file: the
 * formulas stopped by a circular reference keep these values (`CellInput.held`).
 */
export interface XlsxCached {
  sheets: { name: string; cells: Record<string, XlsxCachedCell> }[]
}

/**
 * A stored value: a formula's, with `spill`, the range its array fills (for a legacy array formula, its range), or a
 * cell in an array's range, with `spillOf`, the address of its formula; as in `CellResult`.
 */
export type XlsxCachedCell = Pick<CellResult, 'value' | 'spill' | 'spillOf'>

/** A file read into a workbook. */
export interface XlsxRead {
  /** The workbook, as the engine accepts it. */
  book: Book
  /** The values the file stores for formulas and the cells their arrays spilled into, for comparison with the engine's result. */
  cached: XlsxCached
  issues: XlsxIssue[]
}

/**
 * What `writeXlsx` writes: a workbook and the values of its formulas: a session's `getSnapshot()`, `{ book, result:
 * compute(book) }`, or a file that was read, with the values it stored, `{ book: read.book, result: read.cached }`.
 */
export interface XlsxSnapshot {
  book: BookInput
  result?: ComputeResult | XlsxCached
}

/** Options of `readXlsx`. */
export interface ReadXlsxOptions {
  /**
   * The language of the issue messages: a loaded locale's tag (`'ru-RU'` once `registerLocale(ruRU)` has run) or a
   * locale object (`ruRU`), en-US by default. An issue's code and other fields are the same in every language. A file
   * does not say what language its reader uses, so the host passes the locale the user chose.
   */
  locale?: string | Locale
  /**
   * How large the file's parts may be once inflated: a single part, 512 MiB by default, and all the parts read together,
   * 1 GiB by default. A file that exceeds them is rejected, and so is a part that inflates to more than the file says it
   * holds. A host that reads files uploaded by users should set them lower.
   */
  limits?: XlsxLimits
}

/** The limits `readXlsx` sets on a file's parts once inflated (`ReadXlsxOptions.limits`): sizes in bytes, as whole numbers of at least 1, and a ratio. */
export interface XlsxLimits {
  /**
   * A single part: 536,870,912 (512 MiB) by default. A part is read as one JavaScript string, so a part over about
   * 512 MiB cannot be read whatever the limit.
   */
  entryBytes?: number
  /** All the parts read, together: 1,073,741,824 (1 GiB) by default. */
  totalBytes?: number
  /**
   * How many times its compressed size a part over 16 MiB may inflate to, at least 1: 200 by default. The parts of xlsx
   * files inflate some 5 to 40 times; a zip bomb's, a thousand times.
   */
  ratio?: number
}

/** Options of `writeXlsx`: the language of its issues, which formula values the file stores, and whether it stores formulas or only values. */
export interface WriteXlsxOptions {
  /** The language of the issue messages, as `readXlsx` accepts it (`ReadXlsxOptions.locale`). */
  locale?: string | Locale
  /**
   * The formula values the file stores. `require`: every formula's value and every value of its array, from the result
   * (an `InputError` when one is missing, since then the result is not this workbook's). `omit`: none, and Excel
   * calculates them when it opens the file. `preserve`: those the result has; a formula without one is reported as
   * `noCachedValues`. By default, `require` with a result and `omit` without one, reported as `noCachedValues` when the
   * workbook has formulas.
   */
  values?: 'require' | 'omit' | 'preserve'
  /**
   * `formulas` (the default): formulas, with their values according to `values`. `values`: values only; a formula's
   * cell gets its value from the result (which is then required).
   */
  mode?: 'formulas' | 'values'
}

/** What `writeXlsx` returns: the file, and what the workbook held that the file could not. */
export interface XlsxWritten {
  /** The xlsx file. */
  data: Uint8Array
  issues: XlsxIssue[]
}
