/** Formula text and typed values without a book: a value shown and typed in by a format and a locale, a sheet's name, the references, call and name at a place of formula text, a function's signature, Find's options over text and the order values compare in. */
import type { Locale } from '../locale/types.js'
import type { CellRect, CheckCode, JsonValue, Refusal } from './values.js'
import type { CellInput, CellLink, CellNote, ExternalBook, NumberFormat } from './book.js'
import type { FindCellsOptions } from './reading.js'

/** `displayValue`'s options: where the value is shown. */
export interface DisplayOptions {
  /** The number format; General without it. */
  numberFormat?: NumberFormat
  /** A loaded locale's tag (`LOCALES`), or a locale object, loaded or not; en-US without one. */
  locale?: string | Locale
  /**
   * The column's width as `ColumnInput.width` (several columns' widths added for a merged area). Without it the width
   * is not limited: General shows all the digits it would in a wide column and no number turns into `#`.
   */
  width?: number
  /** The date system a date shows by, as `BookInput.dateSystem`; 1900 without it. */
  dateSystem?: 1900 | 1904
}

/**
 * `entryWidth`'s options: the format the value shows by (the cell's after the entry), the book's locale and date system,
 * and the column typed into, its `width` (the default without one) and `bestFit`.
 */
export interface EntryWidthOptions extends DisplayOptions {
  /** The column's width is one Excel fitted (`ColumnInput.bestFit`), which typing widens as it widens the default width. */
  bestFit?: boolean
}

/** `parseEntry`'s options: the book text is typed into and the cell it goes to. */
export interface EntryOptions {
  /**
   * The book's locale, a loaded locale's tag (`LOCALES`) or a locale object, loaded or not; en-US without one: numbers,
   * dates, booleans, errors and formulas are read as Excel reads them there.
   */
  locale?: string | Locale
  /** The book's date system, as `BookInput.dateSystem`; 1900 without it. */
  dateSystem?: 1900 | 1904
  /**
   * The cell's entry context: its effective `numberFormat`, `style`, `meta`, `note` and `link`; other fields are not read.
   * A `CellInput` or `CellView` of `getCell` or `getCells` carries only the cell's own format. To account for a row's or
   * column's format, pass the resolved cell of `getViewport`, or set `numberFormat` to the effective format yourself.
   * The new input keeps its style, meta, note and link, and its format where Excel keeps it.
   * The returned input stores that format on the cell; `setCellText` keeps an inherited format on its row or column.
   */
  cell?: { numberFormat?: NumberFormat; style?: number; meta?: JsonValue; note?: CellNote; link?: CellLink } | null
  /** The moment whose year completes a date typed without one (`15.3` in ru-RU, `3/15` in en-US), ISO as `Clock.now`; the current moment without it. */
  now?: string
  /**
   * The book's links to other workbooks (`BookInput.externalBooks`): a formula's reference to one, typed as Excel shows
   * it (`=[Book2.xlsx]Sheet1!A1`, `='C:\data\[Book2.xlsx]Sheet1'!A1`, `=Book2.xlsx!Rate`) or by number, is found by the
   * file name and numbered (`'[1]Sheet1'!A1`); one to a book they do not hold is refused (`external_ref`). Without them
   * such a formula is kept as typed.
   */
  externalBooks?: ExternalBook[]
}

/**
 * Text typed into a cell as Excel takes it (`parseEntry`): the cell's new input, `null` when nothing is left of the
 * cell; or a formula Excel does not take, with a check code (`syntax`, `too_long`, `too_deep`), a message in the
 * language of `opts.locale` (English for en-US) and the position in the typed text. A formula typed with brackets,
 * quotes or array braces left open at its end Excel closes on Return when it begins with a function, has no `)` and
 * leaves only functions' brackets open (`=SUM(1,2` is `=SUM(1,2)`); any other such formula it refuses and offers to
 * correct, when the formula closed reads with every function known: `correction` is that text, as typed with what was
 * left open closed (`=(1+2` offers `=(1+2)`). Pasted text is not closed.
 */
export type EntryResult = { ok: true; input: CellInput | null } | (Refusal<CheckCode> & { at: number; correction?: string })

/**
 * `entryText`'s and `needsApostrophe`'s options: the cell's number format (its row's or column's without its own), the
 * book's locale and date system, and for text the moment whose year completes a date typed without one, as `parseEntry`
 * reads it.
 */
export interface EntryTextOptions {
  numberFormat?: NumberFormat
  /** A loaded locale's tag (`LOCALES`) or a locale object; en-US without one. */
  locale?: string | Locale
  dateSystem?: 1900 | 1904
  /** ISO, as `Clock.now`; the current moment without it. */
  now?: string
}

/** Why Excel does not take a sheet name: empty, over 31 characters, one of `: \ / ? * [ ]`, an apostrophe at either end, `History`. */
export type SheetNameCode = 'empty' | 'too_long' | 'character' | 'apostrophe' | 'reserved'

/** `validateSheetName`'s answer; `at` is the position of the character Excel does not take. */
export type SheetNameCheck = { ok: true } | Refusal<SheetNameCode>

/** Where a reference stands in formula text (`scanReferences`): from `at` to `end` (exclusive), with its text as written. */
interface ReferenceSpan {
  at: number
  end: number
  text: string
}

/**
 * A reference to cells, whole columns or whole rows, with its rectangle (`CellRect`) and its sheet as written (unquoted):
 * `sheet` alone for one sheet, `sheet` and `sheetLast` for a reference across sheets (`Jan:Mar!A1`), neither for the
 * formula's own sheet. `spill`: a spill reference (`E2#`), whose rectangle is the formula's cell.
 */
export interface GridReference extends ReferenceSpan, CellRect {
  kind: 'cells' | 'columns' | 'rows'
  sheet?: string
  sheetLast?: string
  spill?: true
}

/**
 * A structured reference (`Sales[Qty]`, `[@Qty]`): the table as written (absent inside a table, `[@Qty]`), the column or
 * the first and last of a span, and the specifiers in stored words (`#Headers`, `#This Row` for `@`); a host lays it on
 * the table's range.
 */
export interface TableReference extends ReferenceSpan {
  kind: 'table'
  table?: string
  columns?: [string] | [string, string]
  items?: ('#All' | '#Data' | '#Headers' | '#Totals' | '#This Row')[]
}

/** A reference found in formula text: to cells or lines of a sheet, or to a table. */
export type FormulaReference = GridReference | TableReference

/**
 * What an arrow does at a place of formula text typed over a cell (Excel's Enter mode), as `refAt` says it: `insert`, a
 * reference to the cell the arrow points at goes in at `at` (Point mode), after an `=` of its own with `equals` (at the
 * start of the text, where Excel begins a formula anew: `=A1` becomes `=D3=A1`); `commit`, the text is entered into the cell and
 * the arrow moves the selection, as after a value; `none`, nothing happens (the text cannot take a reference there and is
 * not a formula Excel enters: inside a string, after a character the locale's formulas do not have).
 */
export type PointAction = { action: 'insert'; at: number; equals?: true } | { action: 'commit' } | { action: 'none' }

/**
 * The call a caret stands in, in formula text as it is typed (`callAt`): the innermost function whose brackets hold the
 * caret, and the argument it is in, for a tip of the function's arguments (`functionSignature`).
 */
export interface CallAt {
  /** The function's name as typed (`СуммЕсли`). */
  name: string
  /** Its English name in upper case, as `functionSignature` and stored formulas take it: the locale's name translated (`SUMIF`), another name as it is. */
  fn: string
  /** The argument the caret is in, from 0: the list separators before it at the call's own level. */
  index: number
  /** Where the call's `(` is, and its `)`, or null while it is open. */
  open: number
  close: number | null
}

/**
 * The name of a function being typed at a caret, in formula text as it is typed (`nameAt`), for the list of Formula
 * AutoComplete: where the name begins and ends, and the part of it before the caret.
 */
export interface NameAt {
  from: number
  to: number
  typed: string
}

/** `scanReferences`'s options. */
export interface ScanOptions {
  /** The locale the text is typed in, a loaded locale's tag (`LOCALES`) or a locale object; without it the text is a stored formula. */
  locale?: string | Locale
}
/** `findMatcher`'s and `findReplacer`'s options: Find's, as `findCells` takes them, over a host's own texts. */
export type FindTextOptions = Pick<FindCellsOptions, 'caseSensitive' | 'wildcards' | 'wholeCell'>

/** The options of `compareValues` and `valueComparator`. */
export interface CompareValuesOptions {
  /** Largest first; blanks stay last. */
  descending?: boolean
  /** A loaded locale's tag (`LOCALES`), or a locale object, loaded or not, whose order of text the values take (`ruRU`, Latin before Cyrillic as Excel for Windows sorts); en-US's without one. */
  locale?: string | Locale
}

/**
 * A function's signature for a tip of its arguments while a formula is typed (`functionSignature`), as Excel's tip lists
 * them: `SUM(number1, [number2], ...)`, in ru-RU `СУММ(число1;[число2];...)`.
 */
export interface FunctionSignature {
  /** The function's name in the locale (`СУММ`). */
  name: string
  /** Its English name, as stored formulas, `CallAt.fn` and `functionDescription` take it (`SUM`). */
  fn: string
  /** The arguments in the tip's order, those that repeat listed twice (`number1`, `number2`), the second time optional. */
  args: FunctionArgument[]
  /**
   * How many of the last arguments repeat past the list (the tip's `...`): 1 for SUM, 2 for SUMIFS, 0 when none do. The
   * argument at index `i` past the list is named as the one `repeat` before it, with its number counted on.
   */
  repeat: number
  /** The tip's text, with the locale's list separator. */
  text: string
  /** The function's category, as Excel's Insert Function window lists it (`FUNCTION_CATEGORIES`). */
  category: FunctionCategory
  /** The category's name in the locale (`Math & Trig`, `Математические`). */
  categoryName: string
}

/**
 * A category of functions, as Excel's Insert Function window lists them; `compatibility` holds the old names Excel keeps
 * beside their newer versions (`NORMDIST` beside `NORM.DIST`).
 */
export type FunctionCategory = 'financial' | 'date' | 'math' | 'statistical' | 'lookup' | 'database' | 'text' | 'logical' | 'information' | 'engineering' | 'compatibility'

/** An argument of a function's signature: its name in the locale and whether a call can leave it out. */
export interface FunctionArgument {
  name: string
  optional: boolean
}
