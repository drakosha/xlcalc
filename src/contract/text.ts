/** Formula text and typed values without a book: a value shown and typed in by a format and a locale, a sheet's name, the references, call and name at a place of formula text, a function's signature, Find's options over text and the order values compare in. */
import type { Locale } from '../locale/types.js'
import type { DateSystem } from '../dates.js'
import type { CellRect, CheckCode, JsonValue, Refusal } from './values.js'
import type { CellInput, CellLink, CellNote, ExternalBook, NumberFormat } from './book.js'
import type { FindCellsOptions } from './reading.js'

/** Options of `displayValue`: where the value is displayed. */
export interface DisplayOptions {
  /** The number format; General without it. */
  numberFormat?: NumberFormat
  /** A loaded locale's tag (`LOCALES`), or a locale object, loaded or not; en-US without one. */
  locale?: string | Locale
  /**
   * The column's width as in `ColumnInput.width` (the sum of several columns' widths for a merged area). Without it, the
   * width is not limited: General displays all the digits it would in a wide column, and no number turns into `#`.
   */
  width?: number
  /** The date system used to display dates, as in `BookInput.dateSystem`; 1900 without it. */
  dateSystem?: DateSystem
}

/**
 * Options of `entryWidth`: the format used to display the value (the cell's format after the entry), the workbook's
 * locale and date system, and the `width` (the default without one) and `bestFit` of the column typed into.
 */
export interface EntryWidthOptions extends DisplayOptions {
  /** The column's width was fitted by Excel (`ColumnInput.bestFit`), so typing widens it as it widens the default width. */
  bestFit?: boolean
}

/** Options of `parseEntry`: the workbook that the text is typed into, and the target cell. */
export interface EntryOptions {
  /**
   * The workbook's locale, a loaded locale's tag (`LOCALES`) or a locale object, loaded or not; en-US without one:
   * numbers, dates, booleans, errors and formulas are parsed as Excel parses them in that locale.
   */
  locale?: string | Locale
  /** The workbook's date system, as in `BookInput.dateSystem`; 1900 without it. */
  dateSystem?: DateSystem
  /**
   * The cell's entry context: its effective `numberFormat`, `style`, `meta`, `note` and `link`; other fields are not read.
   * A `CellInput`, or a `CellView` from `getCell` or `getCells`, carries only the cell's own format. To take a row's or
   * column's format into account, pass the resolved cell from `getViewport`, or set `numberFormat` to the effective
   * format yourself. The new input keeps the cell's style, meta, note and link, and its format where Excel keeps it.
   * The returned input stores that format on the cell; `setCellText` keeps an inherited format on its row or column.
   */
  cell?: { numberFormat?: NumberFormat; style?: number; meta?: JsonValue; note?: CellNote; link?: CellLink } | null
  /** The point in time whose year completes a date typed without a year (`15.3` in ru-RU, `3/15` in en-US), in ISO format as in `Clock.now`; the current time without it. */
  now?: string
  /**
   * The workbook's links to other workbooks (`BookInput.externalBooks`): a formula's reference to one of them, typed as
   * Excel displays it (`=[Book2.xlsx]Sheet1!A1`, `='C:\data\[Book2.xlsx]Sheet1'!A1`, `=Book2.xlsx!Rate`) or by number,
   * is matched by file name and replaced by its number (`'[1]Sheet1'!A1`); a reference to a workbook not among the links
   * is rejected (`external_ref`). Without them, such a formula is kept as typed.
   */
  externalBooks?: ExternalBook[]
  /**
   * True: the text was edited within the cell's existing text (F2, a double-click, the formula bar), as `entryText`
   * writes it, rather than typed over the cell; Excel's percent entry then adds `%` only to a number that begins with a
   * sign: in a cell formatted `0%`, `6%` is 0.06 and `6` is 6 (600%). False by default, as `Session.setCellText` has it.
   */
  inPlace?: boolean
  /**
   * The number format of a cell of the workbook as the workbook stores it: its own, else its row's or column's;
   * undefined for General. `sheet` is the sheet's name as the formula writes it, null for the sheet typed into; `cell`
   * is an A1 address (`C3`). Given it, a typed formula in General takes the format of the cells it refers to, as Excel
   * gives it: `=C3+C1` over a percent in C3 shows a percent. The function is called once for each reference to the
   * workbook's own cells, with the first cell of a range (`C1` for `C:C`); a name or another workbook's cell gives no
   * format. Without it, a reference gives no format.
   */
  numberFormatAt?: (sheet: string | null, cell: string) => NumberFormat | undefined
}

/**
 * Text typed into a cell, as Excel interprets it (`parseEntry`): the cell's new input, `null` when nothing is left of the
 * cell; or a formula Excel would reject, with a check code (`syntax`, `empty`, `too_long`, `too_deep`, and `external_ref`
 * for a workbook not among `opts.externalBooks`), a message in the language of `opts.locale` (English for en-US) and the
 * position in the typed text. When a formula is typed with
 * brackets, quotes or array braces left open at its end, Excel closes them on Return if the formula begins with a
 * function, has no `)` and leaves only function brackets open (`=SUM(1,2` becomes `=SUM(1,2)`); any other such formula
 * it rejects and offers to correct, provided the closed formula parses and every function in it is known: `correction`
 * is that text, as typed with everything left open closed (`=(1+2` offers `=(1+2)`). Pasted text is not closed.
 */
export type EntryResult = { ok: true; input: CellInput | null } | (Refusal<CheckCode> & { at: number; correction?: string })

/**
 * Options of `entryText` and `needsApostrophe`: the cell's number format (its row's or column's when it has none of its
 * own), the workbook's locale and date system, and, for text, the point in time whose year completes a date typed
 * without one, as `parseEntry` parses it.
 */
export interface EntryTextOptions {
  numberFormat?: NumberFormat
  /** A loaded locale's tag (`LOCALES`) or a locale object; en-US without one. */
  locale?: string | Locale
  dateSystem?: DateSystem
  /** ISO, as in `Clock.now`; the current time without it. */
  now?: string
}

/** Why Excel rejects a sheet name: empty, longer than 31 characters, containing one of `: \ / ? * [ ]`, an apostrophe at either end, or `History`. */
export type SheetNameCode = 'empty' | 'too_long' | 'character' | 'apostrophe' | 'reserved'

/** The result of `validateSheetName`; `at` is the position of the character Excel rejects. */
export type SheetNameCheck = { ok: true } | Refusal<SheetNameCode>

/** Where a reference stands in formula text (`scanReferences`): from `at` to `end` (exclusive), with its text as written. */
interface ReferenceSpan {
  at: number
  end: number
  text: string
}

/**
 * A reference to cells, entire columns or entire rows, with its rectangle (`CellRect`) and its sheet as written
 * (unquoted): `sheet` alone for a single sheet, `sheet` and `sheetLast` for a reference that spans sheets
 * (`Jan:Mar!A1`), neither for the formula's own sheet. `spill`: a spill reference (`E2#`), whose rectangle is the
 * formula's cell.
 */
export interface GridReference extends ReferenceSpan, CellRect {
  kind: 'cells' | 'columns' | 'rows'
  sheet?: string
  sheetLast?: string
  spill?: true
}

/**
 * A structured reference (`Sales[Qty]`, `[@Qty]`): the table name as written (absent inside a table, `[@Qty]`), the
 * column or the first and last columns of a span, and the item specifiers in their stored form (`#Headers`, `#This Row`
 * for `@`); a host maps it onto the table's range.
 */
export interface TableReference extends ReferenceSpan {
  kind: 'table'
  table?: string
  columns?: [string] | [string, string]
  items?: ('#All' | '#Data' | '#Headers' | '#Totals' | '#This Row')[]
}

/** A reference found in formula text: to cells, rows or columns of a sheet, or to a table. */
export type FormulaReference = GridReference | TableReference

/**
 * What an arrow key does at a position in formula text typed over a cell (Excel's Enter mode), as `refAt` determines
 * it: `insert`, a reference to the cell the arrow points to is inserted at `at` (Point mode), preceded by its own `=`
 * with `equals` (at the start of the text, where Excel starts a new formula: `=A1` becomes `=D3=A1`); `commit`, the text
 * is entered into the cell and the arrow moves the selection, as after a value; `none`, nothing happens (the text cannot
 * take a reference there and is not a formula Excel would enter: inside a string, or after a character that does not
 * occur in the locale's formulas).
 */
export type PointAction = { action: 'insert'; at: number; equals?: true } | { action: 'commit' } | { action: 'none' }

/**
 * The function call containing the caret in formula text being typed (`callAt`): the innermost function whose brackets
 * enclose the caret, and the argument the caret is in, for a tooltip of the function's arguments (`functionSignature`).
 */
export interface CallAt {
  /** The function's name as typed (`СуммЕсли`). */
  name: string
  /** Its English name in uppercase, as `functionSignature` and stored formulas use it: a locale's name translated (`SUMIF`), any other name as is. */
  fn: string
  /** The argument the caret is in, from 0: the number of list separators before it at the call's own level. */
  index: number
  /** The positions of the call's `(` and of its `)`, or null while the call is still open. */
  open: number
  close: number | null
}

/**
 * The name of a function being typed at the caret, in formula text being typed (`nameAt`), for the Formula AutoComplete
 * list: where the name begins and ends, and the part of it before the caret.
 */
export interface NameAt {
  from: number
  to: number
  typed: string
}

/** Options of `scanReferences`. */
export interface ScanOptions {
  /** The locale the text is typed in, a loaded locale's tag (`LOCALES`) or a locale object; without it, the text is treated as a stored formula. */
  locale?: string | Locale
}
/** Options of `findMatcher` and `findReplacer`: Find's options, as `findCells` accepts them, applied to a host's own texts. */
export type FindTextOptions = Pick<FindCellsOptions, 'caseSensitive' | 'wildcards' | 'wholeCell'>

/** Options of `compareValues` and `valueComparator`. */
export interface CompareValuesOptions {
  /** Largest first; blanks stay last. */
  descending?: boolean
  /** A loaded locale's tag (`LOCALES`), or a locale object, loaded or not, whose text sort order is used (`ruRU`, Latin before Cyrillic, as Excel for Windows sorts); en-US's order without one. */
  locale?: string | Locale
}

/**
 * A function's signature for an argument tooltip while a formula is being typed (`functionSignature`), as Excel's
 * tooltip lists the arguments: `SUM(number1, [number2], ...)`, in ru-RU `СУММ(число1;[число2];...)`.
 */
export interface FunctionSignature {
  /** The function's name in the locale (`СУММ`). */
  name: string
  /** Its English name, as stored formulas, `CallAt.fn` and `functionDescription` use it (`SUM`). */
  fn: string
  /** The arguments in the tooltip's order, with repeating ones listed twice (`number1`, `number2`), the second time as optional. */
  args: FunctionArgument[]
  /**
   * How many of the last arguments repeat beyond the list (the tooltip's `...`): 1 for SUM, 2 for SUMIFS, 0 when none
   * do. The argument at index `i` beyond the list is named like the one `repeat` positions before it, with its number
   * incremented.
   */
  repeat: number
  /** The tooltip's text, with the locale's list separator. */
  text: string
  /** The function's category, as Excel's Insert Function dialog lists it (`FUNCTION_CATEGORIES`). */
  category: FunctionCategory
  /** The category's name in the locale (`Math & Trig`, `Математические`). */
  categoryName: string
}

/**
 * A function category, as Excel's Insert Function dialog lists them; `compatibility` holds the old names Excel keeps
 * alongside their newer versions (`NORMDIST` alongside `NORM.DIST`).
 */
export type FunctionCategory = 'financial' | 'date' | 'math' | 'statistical' | 'lookup' | 'database' | 'text' | 'logical' | 'information' | 'engineering' | 'compatibility'

/** An argument in a function's signature: its name in the locale and whether a call can omit it. */
export interface FunctionArgument {
  name: string
  optional: boolean
}
