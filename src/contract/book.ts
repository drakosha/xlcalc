/** A book as the engine takes it: cells with their notes and links, sheets with their columns, rows, looks, frozen panes, outline settings, tables and view, the book's names and the workbooks it links to; a book without its cells. */
import type { CellValue, ErrorCode, JsonValue, Scalar } from './values.js'
import type { CellStyle, ConditionalFormat, DataValidation } from './looks.js'

/** An input cell. A date is passed explicitly (`{ value: { date: '2025-01-31' } }`) and becomes an Excel serial number. */
export interface CellInput {
  value?: Scalar | { date: string }
  /**
   * Formula text without `=`; a single leading `=` is dropped as the book is taken (`=SUM(A1:A2)` is `SUM(A1:A2)`, the
   * same input and `inputHash`; `==1` stays an error). When present, `value` is ignored.
   */
  formula?: string
  /**
   * Number format, an Excel format code as an xlsx file stores it (en-US syntax: `#,##0.00`, `dd.mm.yyyy`, `[Red]`),
   * shown in the book's locale (`CellResult.text`), or a built-in format by its number (`NumberFormat`). Without it
   * a cell shows its row's format, else its column's.
   */
  numberFormat?: NumberFormat
  /** A style of the book's `styles`, by index. Without it a cell shows its row's style, else its column's. */
  style?: number
  /** Arbitrary cell data (for example ids of source rows), passed through untouched. */
  meta?: JsonValue
  /**
   * An array formula of the old kind (entered with Ctrl+Shift+Enter; in xlsx `<f t="array" ref="B2:C5">` without
   * dynamic array metadata), on the top-left cell of its area: the area (`B2:C5`, or the cell alone, `B2`). The
   * formula is computed as an array, with no implicit intersection, and fills the area, which it never leaves:
   * a smaller array is repeated when it has one row or one column, else the cells past it are #N/A, a larger one is
   * cut, one value fills every cell; nothing spills and nothing blocks it. `ROW()` and `COLUMN()` give the rows and
   * columns of the area. The other cells of the area hold no input, and a session refuses to change some of them
   * without the rest (`Session.setCells`). Only with `formula`.
   */
  array?: string
  /** A note on the cell (Excel's comment of the old kind). Like a style, it goes with the cell and takes no part in computation. */
  note?: CellNote
  /** A hyperlink on the cell; the cell shows its value. Like a style, it goes with the cell and takes no part in computation. */
  link?: CellLink
  /**
   * The value a formula holds while a circular reference stops it. In a book without `iteration` a formula in a cycle,
   * or one that reads a cycle (through other cells too), is not computed, as Excel computes none of them: it shows
   * `held`, or 0 without it (a formula entered there shows 0 in Excel). Elsewhere the field has no effect, and the book
   * as the engine takes it leaves it out, so it counts in `inputHash` in such a cell alone. A session keeps it as Excel
   * keeps such cells: a cell that another edit brings into a cycle's way holds what it showed, a formula entered there
   * holds 0, a cell out of the way is computed again and holds nothing; `getBook()` gives it (`docs/cycles.md`). The
   * values a file stores come in as `held` through `cached` (`ComputeOptions.cached`).
   */
  held?: Exclude<CellValue, null>
}

/** A note: its author and its text, lines separated by `\n`. */
export interface CellNote {
  author?: string
  text: string
}

/**
 * A hyperlink, at least one of `url` and `location`. `url`: an address outside the book, as Excel's Address
 * (`https://example.com/`, `mailto:a@example.com`, a file). `location`: a place in the book, as Excel's SubAddress
 * (`'Sheet 2'!B12`, `Data!A1:C3`, a defined name), or with `url` a place in what it points to (the part after `#`).
 */
export interface CellLink {
  url?: string
  location?: string
  /** The text Excel shows when the pointer rests on the cell (ScreenTip). */
  tooltip?: string
}

/**
 * A number format: a format code as an xlsx file stores it, or the number of one of Excel's built-in formats, 0 to
 * 49, as a file stores a cell's format with no code of its own (`numFmtId`). Excel shows some built-in formats by the
 * locale: 14 is the locale's short date (`m/d/yyyy` in en-US, `dd.mm.yyyy` in ru-RU), 5 to 8 and 37 to 44 its money
 * and number formats, 15 to 22 its other dates and times. A code is shown as written whatever the locale.
 */
export type NumberFormat = string | number

/** Columns of a sheet, as xlsx `<col>`: the width, and the style and format of their cells that have none. */
export interface ColumnInput {
  /**
   * In characters of the default font's digit with the cell's padding, as xlsx stores it, 0 to 255. Without it a
   * column has Excel's default width (8 digits and the padding; Excel for Windows shows it as 8.43).
   */
  width?: number
  /**
   * The width is one Excel fitted, by AutoFit or as a number typed into the column widened it, not one set by hand (xlsx
   * `bestFit`): a number typed into the column widens it again (`entryWidth`), as into a column of the default width.
   */
  bestFit?: boolean
  style?: number
  numberFormat?: NumberFormat
  /** Hidden (xlsx `hidden`); a column of width 0 is hidden too. SUBTOTAL and AGGREGATE count the cells of hidden columns all the same, as Excel does. */
  hidden?: boolean
  /** The outline level (Excel's Group, xlsx `outlineLevel`), 1 to 7; a column without one is not grouped. */
  outlineLevel?: number
  /** The mark of a collapsed group on its summary column (xlsx `collapsed`), as `RowInput.collapsed`. */
  collapsed?: boolean
}

/** Rows of a sheet, as xlsx `<row>`: the height, and the style and format of their cells that have none. */
export interface RowInput {
  /** Points, 0 to 409.5. Without it a row has Excel's default height, 15 points (a line of Calibri 11). */
  height?: number
  /**
   * The height is one Excel fitted, by AutoFit or as text that wraps was typed into the row, not one set by hand (xlsx
   * writes it without `customHeight`): text typed into the row fits it again (`fitRows`), as a row of the default height.
   */
  bestFit?: boolean
  style?: number
  numberFormat?: NumberFormat
  /**
   * Hidden (xlsx `hidden`), by hand or by a filter: SUBTOTAL with codes 101–111 and AGGREGATE with options 1, 3, 5
   * and 7 skip the row's cells, and on a `filtered` sheet SUBTOTAL with codes 1–11 too, as Excel does. A row of
   * `height` 0 is hidden too, as Excel takes it. A collapsed group's rows are hidden rows.
   */
  hidden?: boolean
  /** The outline level (Excel's Group, xlsx `outlineLevel`), 1 to 7; a row without one is not grouped. */
  outlineLevel?: number
  /**
   * The mark of a collapsed group on its summary row (xlsx `collapsed`). Excel keeps it on every summary row whose group
   * has all its rows hidden, however they were hidden, and drops it when one shows; a summary row inside a group whose rows
   * are all hidden keeps the mark it had, so an inner group stays collapsed when the outer one is shown again. A book may
   * hold other marks (`setLayout` sets the field as given): `outlineGroups` and the commands read the marks as Excel
   * keeps them, and `writeXlsx` writes them so.
   */
  collapsed?: boolean
}

/** A sheet of a book (`BookInput.sheets`): its name, its cells by address, and what else a sheet keeps (rows, columns, merged areas, tables, looks, rules, frozen panes). */
export interface SheetInput {
  name: string
  /** Cells by A1 address (`B12`). */
  cells: Record<string, CellInput>
  /**
   * The look of the cells that have no entry in `cells`, by rectangles, each over those before it: a sheet formatted
   * across many empty cells (whole rows to XFD) is held without an object per cell. A cell without an entry shows the
   * style and the number format of the last rectangle that covers it, both from it (a rectangle without a field gives
   * the cell none of its own); a cell with an entry shows its own fields, whatever the rectangles say. A field a cell
   * has none of comes from its row, else its column, as for an entry. `[{ range: 'A20:XFD22', style: 0 },
   * { range: 'R20:R22', style: 1, numberFormat: '0.0%' }]` is the same sheet as an entry `{ style: 0 }` in each cell of
   * A20:XFD22 but R20:R22, which hold `{ style: 1, numberFormat: '0.0%' }`. A book is taken with every entry that has
   * only a style or a number format moved here, and the calls that give a book back give this list in a form of their
   * own (`LookRange`).
   */
  looks?: LookRange[]
  /** Columns by letters, one (`C`) or a span (`B:D`) as xlsx `<col min max>`; spans do not overlap. */
  columns?: Record<string, ColumnInput>
  /** Rows by number, one (`5`) or a span (`5:9`); spans do not overlap. */
  rows?: Record<string, RowInput>
  /** Merged areas (`B2:D3`), as xlsx `<mergeCell>`: two cells or more each, none overlapping another. */
  merges?: string[]
  /**
   * A filter of the sheet has criteria, its autofilter's or a table's (Excel's filter mode): every hidden row of the
   * sheet counts as one a filter hid, which SUBTOTAL skips with codes 1–11 too and a copy leaves out, as Excel does
   * whatever hid the row.
   */
  filtered?: boolean
  /** The sheet's tables (Excel's list objects); they do not overlap, and no merged area or array area of several cells meets one. */
  tables?: TableInput[]
  /**
   * Frozen panes (Excel's Freeze Panes): the first `rows` rows and the first `columns` columns stay in view as the sheet
   * scrolls, one of them at least (`{ rows: 1 }` freezes the top row). Not part of computation.
   */
  frozen?: FrozenPanes
  /** Saved window state from xlsx; does not affect computation or inputHash. */
  view?: SheetView
  /**
   * The sheet's conditional formats (Excel's conditional formatting rules) in priority order, the first one first. They
   * do not change values: a session gives what they show in `CellView.appearance` (`docs/conditional-formats.md`).
   */
  conditionalFormats?: ConditionalFormat[]
  /**
   * Data validation (Excel's Data Validation): the rules of the sheet's cells, at most one rule over a cell. Not part of
   * computation: a session gives a cell's rule and its list (`Session.getValidation`) and checks a value against it
   * (`Session.validateEntry`).
   */
  validations?: DataValidation[]
  /** Where the summary rows and columns of the sheet's groups are (`RowInput.outlineLevel`); below and right by default. */
  outline?: OutlineSettings
  /**
   * The sheet is hidden (Excel's Hide, xlsx `state="hidden"`), or very hidden (`'very'`, xlsx `state="veryHidden"`, a
   * sheet Excel's window does not offer to show again, only its object model). Not part of computation: formulas read a
   * hidden sheet as any other. A book keeps one visible sheet at least.
   */
  hidden?: boolean | 'very'
}

/**
 * A rectangle of a sheet's looks (`SheetInput.looks`): a cell (`B2`), a range (`B2:D9`), whole columns (`B:D`) or whole
 * rows (`2:9`), and the style and the number format its cells without an entry show; neither field, and its cells show
 * none of their own. The books the engine gives back hold the looks in one form for one set of them: rows in blocks, the
 * runs of a group of rows as rectangles across the group, then the rows of the group that differ, its rows without looks
 * as rows of none.
 */
export interface LookRange {
  range: string
  /** A style of the book's `styles`, by index. */
  style?: number
  numberFormat?: NumberFormat
}

/** A sheet's frozen panes (`SheetInput.frozen`, Excel's Freeze Panes): the rows above the split and the columns left of it. */
export interface FrozenPanes {
  /** 1 to 1,048,575. */
  rows?: number
  /** 1 to 16,383. */
  columns?: number
}

/** The summary of a group of rows is the row below it (`summaryBelow`, true by default) or above; of columns, right of it or left. */
export interface OutlineSettings {
  summaryBelow?: boolean
  summaryRight?: boolean
}

/**
 * A table (Excel's list object): a rectangle of a sheet with a header row of column names, data rows and a totals row.
 * Formulas refer to its parts by name with structured references (`Sales[Amount]`, `SUM(Sales[[#Totals],[Qty]])`),
 * which follow the table as rows and columns are inserted and deleted in it. A formula stores a reference to its own
 * row as an xlsx file does, `Sales[[#This Row],[Amount]]`; `Sales[@Amount]`, and `[@Amount]` inside the table, are read too.
 */
export interface TableInput {
  /**
   * Excel's rules for names: a letter of any script, `_` or `\` first, then letters, digits, `_`, `.`, `\` and `?`, at
   * most 255 characters; not a cell address (`T2`) or R1C1 (`R`, `C`, `R1C1`), not TRUE or FALSE, and not the book's
   * locale's words for them either, in any case (ИСТИНА and ЛОЖЬ in ru-RU, as Excel for Windows ru-RU refuses them).
   * Tables and defined names share one namespace: no other table and no defined name of the book has the name, in any case.
   */
  name: string
  /** The table's rectangle, its header and totals rows included (`B2:E10`), with one data row at least. */
  range: string
  /** The first row is a header row (true by default): its cells hold the columns' names as text, and no formulas. */
  headers?: boolean
  /** The last row is a totals row (false by default). */
  totals?: boolean
  /** The columns from left to right, one for each column of the range. */
  columns: TableColumn[]
}

/** A column of a table (`TableInput.columns`): its header's name and what its cells and totals row compute. */
export interface TableColumn {
  /** 1 to 255 characters, unique in the table in any case; the header cell's text. */
  name: string
  /**
   * The formula of a calculated column (stored text, `Sales[[#This Row],[Qty]]*2`; a leading `=` is dropped), which a row
   * inserted into the table gets in this column, as Excel fills it. The column's cells hold their own formulas.
   */
  formula?: string
  /**
   * What the column's cell of the totals row is, as an xlsx file records it: one of Excel's totals functions, a
   * SUBTOTAL over the column that the cell holds as its formula (`custom`: a formula of its own), or a label. The cell
   * computes what it holds.
   */
  totals?: { function: TotalsFunction } | { label: string }
}

/** Excel's totals functions of a table column (xlsx `totalsRowFunction`). */
export type TotalsFunction = 'sum' | 'average' | 'count' | 'countNums' | 'max' | 'min' | 'stdDev' | 'var' | 'custom'

/** A defined name of a book (`BookInput.names`): the name, what it refers to and its scope. */
export interface NameInput {
  /**
   * Excel's rules for names, a table's too: a letter of any script, `_` or `\` first, then letters, digits, `_`, `.`,
   * `\` and `?`, at most 255 characters; not a cell address (`Q1`) or R1C1 (`R`, `C`, `R2C3`), not TRUE or FALSE, and
   * not the book's locale's words for them either, in any case (ИСТИНА and ЛОЖЬ in ru-RU, which Excel ru-RU refuses;
   * `Ставка`, `НДС_2025`, `АБ12` and `Став?ка` are names). Compared in any case: `ставка` is the name `Ставка`.
   */
  name: string
  /**
   * What the name refers to: a formula without `=` (a leading one is dropped), as Excel's "Refers to" and xlsx
   * `<definedName>` hold it. A reference with a sheet (`'Data'!$D$2:$D$3401`, `'Data'!$D:$D`), a constant (`0.2`,
   * `"text"`, `{1,2,3}`), a formula (`'Data'!$A$1*2`, `SUM('Data'!$A:$A)`, `INDEX('Data'!$A$1:$C$9,2,0)`, which is a
   * reference), a union of references on one sheet (`'Data'!$A$1:$A$3,'Data'!$C$1:$C$3`). Every reference has its sheet.
   * A reference without `$` is relative to A1, as a file stores it: used in a cell it moves by that cell's offset from
   * A1, wrapping round the sheet's edges (`'Data'!B1` used in C5 is `'Data'!D5`, `'Data'!A1048576` is the row above).
   * Names in it are the global ones (a sheet's own as `'Sheet'!name`), as a file stores a name's formula. The formula is
   * computed where the name is used, as part of that formula.
   */
  ref: string
  /** The sheet where the name is visible; a local name shadows a global one. Without it the name is global. */
  scope?: string
}

/** The first saved sheet view; selection holds A1 cells or ranges, without a sheet name. */
export interface SheetView {
  activeCell?: string
  selection?: string[]
  topLeft?: string
  /** Zoom in percent, 10 to 400. */
  zoom?: number
}

/** A workbook as the engine takes it: its sheets with their cells, names, styles, locale and date system; what `compute`, `createSession` and `writeXlsx` take and `readXlsx` gives. */
export interface BookInput {
  /** The saved active visible sheet, by name; the writer defaults to the first visible sheet. Not part of inputHash. */
  activeSheet?: string
  /** Excel's rules for sheet names: 1 to 31 characters, none of `: \ / ? * [ ]`, no apostrophe at either end; one visible sheet at least. */
  sheets: SheetInput[]
  names?: NameInput[]
  /** The cell styles that cells, rows and columns refer to by index. */
  styles?: CellStyle[]
  /**
   * The locale the book is computed in, the BCP 47 tag of a loaded locale, `LOCALES` (`'en-US'` by default; another
   * of the package's locales is loaded by `registerLocale`). It changes computed values as Excel's regional settings do
   * (text read as a number or a date, numbers and booleans turned into text, `TEXT` format codes, currency), not the
   * formula text: formulas stay invariant, and `toLocal` shows them in a locale. A tag that is not loaded is an
   * `InputError` that names the module to import.
   */
  locale?: string
  /**
   * The date system, as xlsx `workbookPr date1904`: 1900 (the default) counts days from 1900-01-00 with Excel's
   * phantom 1900-02-29, 1904 from 1904-01-01, serial 0. It changes the serial numbers of the dates typed in and
   * computed, TODAY and NOW, date text and how dates show: a 1904 book shows a negative date or time with a minus
   * (`-12:00`). It goes into both hashes only when it is 1904.
   */
  dateSystem?: 1900 | 1904
  /**
   * Iterative calculation of circular references, as Excel's setting and xlsx `calcPr iterate`. Without it the cells of a
   * cycle and the cells that read it are not computed and hold their values (`CellInput.held`). With it the cells of
   * every cycle and the cells that depend on them are computed from 0, again and again, sheet by sheet in the order of
   * the sheets' names (in any case), row by row. That stops after `maxIterations` rounds (100 by default, 1 to 32,767)
   * or when no value changes by `maxChange` or more (0.001 by default). Each computation starts from 0 again, so a
   * session gives what a full computation gives, and `held` has no effect. Excel goes on from the values it holds
   * instead.
   */
  iteration?: { maxIterations?: number; maxChange?: number }
  /**
   * The other workbooks the book's formulas refer to (`[1]Sheet1!A1`, `'[1]My Sheet'!A1:B9`, `[1]!Rate`): element `i`
   * answers `[i+1]`. They are computed from the values stored here, as Excel computes a file it opens without updating
   * its links; the other book itself is never read. It goes into `inputHash` when it is not empty.
   */
  externalBooks?: ExternalBook[]
}

/**
 * Another workbook as a book keeps it beside the formulas that read it (xlsx `externalLinkN.xml`): where it is, and
 * the values Excel stored of it when it last read it.
 */
export interface ExternalBook {
  /** Where the link points, as the file holds it: a file name, a relative or an absolute path, a `file:` or web address. */
  target: string
  /** Its sheets in the order the file lists them. */
  sheets: ExternalSheet[]
  /** Its defined names the book's formulas use: of the book, or of one of its sheets. */
  names?: ExternalName[]
}

/** A sheet of another workbook and the values stored of it. */
export interface ExternalSheet {
  name: string
  /**
   * The stored values by A1 address, null for a cell stored empty. A cell they do not hold is blank, or `#REF!` with
   * `refreshError`. Without `cells` nothing of the sheet is stored, and every cell of it is `#REF!`.
   */
  cells?: Record<string, ExternalValue | null>
  /**
   * Excel's mark that the last update of the sheet's values failed (`refreshError`): a cell they do not hold is `#REF!`
   * instead of blank. Excel sets it on every sheet of a closed book when a formula referring to that book is entered,
   * and so does a session.
   */
  refreshError?: true
}

/** A value stored of another workbook's cell. */
export type ExternalValue = number | string | boolean | { error: ErrorCode }

/** A defined name of another workbook. */
export interface ExternalName {
  name: string
  /** What it refers to, as the file holds it without `=`: a reference into its book (`Sheet1!$B$1`), `#REF!`. */
  ref: string
  /** The index in `sheets` of the sheet whose own name it is; a name of the book has none. */
  sheet?: number
}

/**
 * A book as the calls that give one answer it (`Session.getBook`, `getSnapshot`, `getLayout`, `editStructure`,
 * `readXlsx`): a `BookInput`, which goes back into every call that takes a book.
 */
export type Book = BookInput

/** A sheet of a `Book`. */
export type Sheet = SheetInput

/** A cell of a `Book`. */
export type Cell = CellInput

/** A sheet without its cells and their looks (`Session.getLayout`). */
export type SheetLayout = Omit<Sheet, 'cells' | 'looks'>

/** A book without the cells of its sheets (`Session.getLayout`): its styles, names and settings, and each sheet's layout. */
export type BookLayout = Omit<Book, 'sheets'> & { sheets: SheetLayout[] }
