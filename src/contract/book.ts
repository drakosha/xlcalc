/** A book as the engine takes it: cells with their notes and links, sheets with their columns, rows, looks, frozen panes, outline settings, tables and view, the book's names and the workbooks it links to; a book without its cells. */
import type { DateSystem } from '../dates.js'
import type { CellValue, ErrorCode, JsonValue, Scalar } from './values.js'
import type { CellStyle, ConditionalFormat, DataValidation } from './looks.js'

/**
 * An input cell. A date is passed explicitly (`{ value: { date: '2025-01-31' } }`) and becomes an Excel serial number; a
 * number must be finite (NaN and infinities throw `InputError`).
 */
export interface CellInput {
  value?: Scalar | { date: string }
  /**
   * Formula text without `=`; a single leading `=` is removed on input (`=SUM(A1:A2)` becomes `SUM(A1:A2)`, with the
   * same input and `inputHash`; `==1` remains an error). When present, `value` is ignored.
   */
  formula?: string
  /**
   * Number format: an Excel format code as stored in an xlsx file (en-US syntax: `#,##0.00`, `dd.mm.yyyy`, `[Red]`),
   * displayed in the workbook's locale (`CellResult.text`), or the number of a built-in format (`NumberFormat`).
   * Without it, a cell uses its row's format, or else its column's.
   */
  numberFormat?: NumberFormat
  /** An index into the workbook's `styles`. Without it, a cell uses its row's style, or else its column's. */
  style?: number
  /** Arbitrary cell data (for example ids of source rows), passed through untouched. */
  meta?: JsonValue
  /**
   * A legacy array formula (entered with Ctrl+Shift+Enter; in xlsx, `<f t="array" ref="B2:C5">` without dynamic array
   * metadata), set on the top-left cell of its range: the range (`B2:C5`, or just the cell, `B2`). The formula is
   * calculated as an array, without implicit intersection, and fills its range without ever extending past it: a
   * smaller array is repeated when it has one row or one column, otherwise the cells beyond it are #N/A; a larger one
   * is truncated; a single value fills every cell. Nothing spills, and nothing blocks it. `ROW()` and `COLUMN()` return
   * the rows and columns of the range. The other cells of the range hold no input, and a session does not allow
   * changing some of them without the rest (`Session.setCells`). Only with `formula`.
   */
  array?: string
  /** A note on the cell (what Excel used to call a comment). Like a style, it stays with the cell and plays no part in calculation. */
  note?: CellNote
  /** A hyperlink on the cell; the cell displays its value. Like a style, it stays with the cell and plays no part in calculation. */
  link?: CellLink
  /**
   * The value a formula keeps while a circular reference prevents its calculation. In a workbook without `iteration`, a
   * formula in a cycle, or one that depends on a cycle (including through other cells), is not calculated, just as Excel
   * calculates none of them: it displays `held`, or 0 without it (a formula entered there displays 0 in Excel).
   * Elsewhere the field has no effect and is dropped on input, so it counts toward `inputHash` only in such cells. A
   * session maintains it the way Excel treats such cells: a cell that another edit puts in a cycle's path keeps the
   * value it displayed, a formula entered there keeps 0, and a cell no longer in the path is recalculated and keeps
   * nothing; `getBook()` returns it (`docs/cycles.md`). The values stored in a file are loaded as `held` through
   * `cached` (`ComputeOptions.cached`).
   */
  held?: Exclude<CellValue, null>
}

/** A note: its author and its text, lines separated by `\n`. */
export interface CellNote {
  author?: string
  text: string
}

/**
 * A hyperlink, with at least one of `url` and `location`. `url`: an address outside the workbook, like Excel's Address
 * (`https://example.com/`, `mailto:a@example.com`, a file). `location`: a place in the workbook, like Excel's
 * SubAddress (`'Sheet 2'!B12`, `Data!A1:C3`, a defined name), or, together with `url`, a place within the target (the
 * part after `#`).
 */
export interface CellLink {
  url?: string
  location?: string
  /** The text Excel displays when the pointer hovers over the cell (ScreenTip). */
  tooltip?: string
}

/**
 * A number format: a format code as stored in an xlsx file, or the number of one of Excel's built-in formats, 0 to
 * 49, which is how a file stores the format of a cell that has no code of its own (`numFmtId`). Excel displays some
 * built-in formats according to the locale: 14 is the locale's short date (`m/d/yyyy` in en-US, `dd.mm.yyyy` in
 * ru-RU), 5 to 8 and 37 to 44 its currency and number formats, 15 to 22 its other date and time formats. A format code
 * is displayed as written, regardless of the locale.
 */
export type NumberFormat = string | number

/** Columns of a sheet, like xlsx `<col>`: the width, and the style and number format for their cells that have none of their own. */
export interface ColumnInput {
  /**
   * In widths of the default font's digit, plus the cell padding, as xlsx stores it, 0 to 255. Without it, a column has
   * Excel's default width (8 digits plus padding; Excel for Windows displays it as 8.43).
   */
  width?: number
  /**
   * The width was set by Excel, by AutoFit or by widening the column for a number typed into it, rather than manually
   * (xlsx `bestFit`): a number typed into the column widens it again (`entryWidth`), as in a column of the default width.
   */
  bestFit?: boolean
  style?: number
  numberFormat?: NumberFormat
  /** Hidden (xlsx `hidden`); a column of width 0 is also hidden. SUBTOTAL and AGGREGATE still include the cells of hidden columns, as Excel does. */
  hidden?: boolean
  /** The outline level (Excel's Group, xlsx `outlineLevel`), 1 to 7; a column without one is not grouped. */
  outlineLevel?: number
  /** The collapsed marker on a group's summary column (xlsx `collapsed`), like `RowInput.collapsed`. */
  collapsed?: boolean
}

/** Rows of a sheet, like xlsx `<row>`: the height, and the style and number format for their cells that have none of their own. */
export interface RowInput {
  /** In points, 0 to 409.5. Without it, a row has Excel's default height, 15 points (one line of Calibri 11). */
  height?: number
  /**
   * The height was set by Excel, by AutoFit or when wrapping text was typed into the row, rather than manually (xlsx
   * writes it without `customHeight`): text typed into the row fits it again (`fitRows`), as for a row of the default
   * height.
   */
  bestFit?: boolean
  style?: number
  numberFormat?: NumberFormat
  /**
   * Hidden (xlsx `hidden`), manually or by a filter: SUBTOTAL with codes 101–111 and AGGREGATE with options 1, 3, 5
   * and 7 ignore the row's cells, and on a `filtered` sheet so does SUBTOTAL with codes 1–11, as in Excel. A row with
   * `height` 0 is also hidden, as Excel treats it. The rows of a collapsed group are hidden rows.
   */
  hidden?: boolean
  /** The outline level (Excel's Group, xlsx `outlineLevel`), 1 to 7; a row without one is not grouped. */
  outlineLevel?: number
  /**
   * The collapsed marker on a group's summary row (xlsx `collapsed`). Excel sets it on every summary row whose group has
   * all its rows hidden, however they were hidden, and clears it when one of them is shown; a summary row inside a group
   * whose rows are all hidden keeps the marker it had, so an inner group stays collapsed when the outer one is expanded
   * again. A workbook may hold other markers (`setLayout` sets the field as given): `outlineGroups` and the outline
   * commands read the markers as Excel would keep them, and `writeXlsx` writes them that way.
   */
  collapsed?: boolean
}

/** A sheet of a workbook (`BookInput.sheets`): its name, its cells by address, and everything else a sheet holds (rows, columns, merged areas, tables, formatting, rules, frozen panes). */
export interface SheetInput {
  name: string
  /** Cells by A1 address (`B12`). */
  cells: Record<string, CellInput>
  /**
   * The formatting of the cells that have no entry in `cells`, as rectangles, each applied over those before it: a sheet
   * formatted across many empty cells (entire rows out to XFD) is stored without an object per cell. A cell without an
   * entry uses the style and the number format of the last rectangle that covers it, both from that rectangle (a
   * rectangle without a field gives the cell none of its own); a cell with an entry uses its own fields, whatever the
   * rectangles say. A field the cell lacks comes from its row, or else its column, as for an entry.
   * `[{ range: 'A20:XFD22', style: 0 }, { range: 'R20:R22', style: 1, numberFormat: '0.0%' }]` is the same sheet as an
   * entry `{ style: 0 }` in each cell of A20:XFD22 except R20:R22, which hold `{ style: 1, numberFormat: '0.0%' }`. On
   * input, every entry that has only a style or a number format is moved here, and the calls that return a workbook
   * return this list in a form of their own (`LookRange`).
   */
  looks?: LookRange[]
  /** Columns by letter, a single column (`C`) or a span (`B:D`) like xlsx `<col min max>`; spans do not overlap. */
  columns?: Record<string, ColumnInput>
  /** Rows by number, a single row (`5`) or a span (`5:9`); spans do not overlap. */
  rows?: Record<string, RowInput>
  /** Merged areas (`B2:D3`), like xlsx `<mergeCell>`: two or more cells each, none overlapping another. */
  merges?: string[]
  /**
   * A filter on the sheet has criteria, the sheet's AutoFilter or a table's (Excel's filter mode): every hidden row of
   * the sheet counts as filtered out, so SUBTOTAL ignores it with codes 1–11 too and a copy skips it, as Excel does
   * whatever hid the row.
   */
  filtered?: boolean
  /** The sheet's tables (Excel's list objects); they do not overlap, and no merged area or multi-cell array range overlaps one. */
  tables?: TableInput[]
  /**
   * Frozen panes (Excel's Freeze Panes): the first `rows` rows and the first `columns` columns stay visible while the
   * sheet scrolls; at least one of the two is set (`{ rows: 1 }` freezes the top row). Not part of calculation.
   */
  frozen?: FrozenPanes
  /** The saved sheet view from xlsx; it does not affect calculation or inputHash. */
  view?: SheetView
  /**
   * The sheet's conditional formats (Excel's conditional formatting rules) in priority order, highest first. They do not
   * change values: a session returns their effect on display in `CellView.appearance` (`docs/conditional-formats.md`).
   */
  conditionalFormats?: ConditionalFormat[]
  /**
   * Data validation (Excel's Data Validation): the rules for the sheet's cells, at most one rule per cell. Not part of
   * calculation: a session returns a cell's rule and its list (`Session.getValidation`) and checks a value against it
   * (`Session.validateEntry`).
   */
  validations?: DataValidation[]
  /** Where the summary rows and columns of the sheet's groups are placed (`RowInput.outlineLevel`); below and to the right by default. */
  outline?: OutlineSettings
  /**
   * The sheet is hidden (Excel's Hide, xlsx `state="hidden"`) or very hidden (`'very'`, xlsx `state="veryHidden"`, a
   * sheet that Excel's interface does not offer to unhide; only its object model can). Not part of calculation: formulas
   * read a hidden sheet like any other. A workbook always has at least one visible sheet.
   */
  hidden?: boolean | 'very'
  /**
   * The sheet is protected (Excel's Protect Sheet, xlsx `<sheetProtection>`), with what users may still do: its cells
   * that are not locked take edits (`CellStyle.protection`). Not part of calculation or `inputHash`, and a session does
   * not enforce it: the host does. The workbook holds no password.
   */
  protection?: SheetProtection
}

/**
 * A rectangle of a sheet's formatting (`SheetInput.looks`): a cell (`B2`), a range (`B2:D9`), entire columns (`B:D`)
 * or entire rows (`2:9`), and the style and number format used by its cells that have no entry; with neither field,
 * its cells have none of their own. The workbooks the engine returns hold the same formatting in a single canonical
 * form: rows in blocks, the runs shared by a group of rows as rectangles across the group, then the rows of the group
 * that differ, and its rows without formatting as rectangles with neither field.
 */
export interface LookRange {
  range: string
  /** An index into the workbook's `styles`. */
  style?: number
  numberFormat?: NumberFormat
}

/** A sheet's frozen panes (`SheetInput.frozen`, Excel's Freeze Panes): the rows above the split and the columns to the left of it. */
export interface FrozenPanes {
  /** 1 to 1,048,575. */
  rows?: number
  /** 1 to 16,383. */
  columns?: number
}

/**
 * What users may do on a protected sheet (`SheetInput.protection`), as the boxes of Excel's Protect Sheet dialog: each
 * field true allows it. Selecting locked and unlocked cells is allowed by default, the rest is not, and a protection
 * holds only what differs from that: `{}` is Excel's default protection, `{ formatCells: true }` also allows
 * formatting cells, `{ selectLockedCells: false }` allows selecting only the unlocked ones.
 */
export interface SheetProtection {
  selectLockedCells?: boolean
  selectUnlockedCells?: boolean
  formatCells?: boolean
  formatColumns?: boolean
  formatRows?: boolean
  insertColumns?: boolean
  insertRows?: boolean
  insertHyperlinks?: boolean
  deleteColumns?: boolean
  deleteRows?: boolean
  sort?: boolean
  /** Use AutoFilter. */
  autoFilter?: boolean
  /** Use PivotTable and PivotChart. */
  pivotTables?: boolean
  /** Edit objects: drawings, charts, notes. */
  objects?: boolean
  /** Edit scenarios. */
  scenarios?: boolean
}

/** A group of rows has its summary row below it (`summaryBelow`, true by default) or above it; a group of columns, to the right or to the left. */
export interface OutlineSettings {
  summaryBelow?: boolean
  summaryRight?: boolean
}

/**
 * A table (Excel's list object): a rectangle on a sheet with a header row of column names, data rows and a totals row.
 * Formulas refer to its parts by name with structured references (`Sales[Amount]`, `SUM(Sales[[#Totals],[Qty]])`),
 * which follow the table as rows and columns are inserted into it and deleted from it. A formula stores a reference to
 * its own row the way an xlsx file does, `Sales[[#This Row],[Amount]]`; `Sales[@Amount]`, and `[@Amount]` inside the
 * table, are also accepted.
 */
export interface TableInput {
  /**
   * Excel's rules for names: a letter of any script, `_` or `\` first, then letters, digits, `_`, `.`, `\` and `?`, at
   * most 255 characters; not a cell address (`T2`) or an R1C1 reference (`R`, `C`, `R1C1`), not TRUE or FALSE, nor the
   * workbook locale's words for them, in any case (ИСТИНА and ЛОЖЬ in ru-RU, which Excel for Windows ru-RU rejects).
   * Tables and defined names share one namespace: no other table and no defined name in the workbook may have the same
   * name, compared case-insensitively.
   */
  name: string
  /** The table's rectangle, including its header and totals rows (`B2:E10`), with at least one data row. */
  range: string
  /** The first row is a header row (true by default): its cells hold the column names as text, not formulas. */
  headers?: boolean
  /** The last row is a totals row (false by default). */
  totals?: boolean
  /** The columns from left to right, one for each column of the range. */
  columns: TableColumn[]
  /**
   * How Excel draws the table (xlsx `tableStyleInfo`); without it, Excel's default, `TableStyleMedium2` with banded rows.
   * The engine does not draw it: it stays with the table through operations and is written to the file.
   */
  style?: TableStyle
}

/**
 * A table's style (`TableInput.style`): one of Excel's table styles, and the parts of the table it emphasizes, each
 * false by default, as the checkboxes of Excel's Table Style Options.
 */
export interface TableStyle {
  /** One of Excel's table styles, `TableStyleLight1` to `TableStyleLight21`, `TableStyleMedium1` to `TableStyleMedium28`, `TableStyleDark1` to `TableStyleDark11`; none without it. */
  name?: string
  firstColumn?: boolean
  lastColumn?: boolean
  /** Banded rows. */
  rowStripes?: boolean
  /** Banded columns. */
  columnStripes?: boolean
}

/** A column of a table (`TableInput.columns`): its header name and what its cells and its totals row calculate. */
export interface TableColumn {
  /** 1 to 255 characters, unique within the table, compared case-insensitively; the header cell's text. */
  name: string
  /**
   * The formula of a calculated column (as stored, `Sales[[#This Row],[Qty]]*2`; a leading `=` is removed), which a row
   * inserted into the table receives in this column, as Excel fills it. The column's cells hold their own formulas.
   */
  formula?: string
  /**
   * What the column's cell in the totals row contains, as an xlsx file records it: one of Excel's totals functions, a
   * SUBTOTAL over the column that the cell holds as its formula (`custom`: a formula of its own), or a label. The cell
   * calculates whatever it holds.
   */
  totals?: { function: TotalsFunction } | { label: string }
}

/** Excel's totals functions of a table column (xlsx `totalsRowFunction`). */
export type TotalsFunction = 'sum' | 'average' | 'count' | 'countNums' | 'max' | 'min' | 'stdDev' | 'var' | 'custom'

/** A defined name of a workbook (`BookInput.names`): the name, what it refers to and its scope. */
export interface NameInput {
  /**
   * Excel's rules for names, the same as for tables: a letter of any script, `_` or `\` first, then letters, digits,
   * `_`, `.`, `\` and `?`, at most 255 characters; not a cell address (`Q1`) or an R1C1 reference (`R`, `C`, `R2C3`),
   * not TRUE or FALSE, nor the workbook locale's words for them, in any case (ИСТИНА and ЛОЖЬ in ru-RU, which Excel
   * ru-RU rejects; `Ставка`, `НДС_2025`, `АБ12` and `Став?ка` are valid names). Names are compared case-insensitively:
   * `ставка` is the name `Ставка`.
   */
  name: string
  /**
   * What the name refers to: a formula without `=` (a leading one is removed), as Excel's "Refers to" box and xlsx
   * `<definedName>` hold it. A reference with a sheet name (`'Data'!$D$2:$D$3401`, `'Data'!$D:$D`), a constant (`0.2`,
   * `"text"`, `{1,2,3}`), a formula (`'Data'!$A$1*2`, `SUM('Data'!$A:$A)`, `INDEX('Data'!$A$1:$C$9,2,0)`, which returns
   * a reference), or a union of references on one sheet (`'Data'!$A$1:$A$3,'Data'!$C$1:$C$3`). Every reference includes
   * its sheet. A reference without `$` is relative to A1, as a file stores it: used in a cell, it shifts by that cell's
   * offset from A1, wrapping around the sheet's edges (`'Data'!B1` used in C5 is `'Data'!D5`, and `'Data'!A1048576` is
   * the row above). Names in it are global names (a sheet-level name is written as `'Sheet'!name`), as a file stores a
   * name's formula. The formula is calculated where the name is used, as part of the formula that uses it.
   */
  ref: string
  /** The sheet where the name is visible; a local name shadows a global one. Without it the name is global. */
  scope?: string
  /**
   * Hidden (xlsx `hidden="1"`): Excel's Name Manager does not list the name, which a host's list of names should follow.
   * Not part of calculation: formulas use a hidden name like any other.
   */
  hidden?: boolean
}

/** The first saved sheet view; selection holds A1 cells or ranges, without a sheet name. */
export interface SheetView {
  activeCell?: string
  selection?: string[]
  topLeft?: string
  /** Zoom in percent, 10 to 400. */
  zoom?: number
}

/** A workbook as the engine accepts it: its sheets with their cells, names, styles, locale and date system; what `compute`, `createSession` and `writeXlsx` accept and `readXlsx` returns. */
export interface BookInput {
  /** The saved active visible sheet, by name; the writer defaults to the first visible sheet. Not part of inputHash. */
  activeSheet?: string
  /** Excel's rules for sheet names: 1 to 31 characters, none of `: \ / ? * [ ]`, no apostrophe at either end; at least one visible sheet. */
  sheets: SheetInput[]
  names?: NameInput[]
  /** The cell styles that cells, rows and columns refer to by index. */
  styles?: CellStyle[]
  /**
   * The locale the workbook is calculated in: the BCP 47 tag of a loaded locale, `LOCALES` (`'en-US'` by default; the
   * package's other locales are loaded with `registerLocale`). It changes calculated values as Excel's regional settings
   * do (text parsed as a number or a date, numbers and booleans converted to text, `TEXT` format codes, currency), but
   * not the formula text: formulas stay invariant, and `toLocal` displays them in a locale. A tag that is not loaded
   * throws an `InputError` that names the module to import.
   */
  locale?: string
  /**
   * The date system, like xlsx `workbookPr date1904`: 1900 (the default) counts days from 1900-01-00, including Excel's
   * phantom 1900-02-29; 1904 counts from 1904-01-01 as serial 0. It changes the serial numbers of typed and calculated
   * dates, TODAY and NOW, date text and how dates are displayed: a 1904 workbook displays a negative date or time with a
   * minus sign (`-12:00`). It is included in both hashes only when it is 1904.
   */
  dateSystem?: DateSystem
  /**
   * Iterative calculation of circular references, like Excel's setting and xlsx `calcPr iterate`. Without it, the cells
   * of a cycle and the cells that depend on it are not calculated and keep their values (`CellInput.held`). With it, the
   * cells of every cycle and the cells that depend on them are calculated from 0, repeatedly, sheet by sheet in order of
   * sheet name (case-insensitively), row by row. Iteration stops after `maxIterations` rounds (100 by default, 1 to
   * 32,767) or when no value changes by `maxChange` or more (0.001 by default). Each calculation starts again from 0, so
   * a session returns what a full calculation returns, and `held` has no effect. Excel instead continues from the values
   * it holds.
   */
  iteration?: { maxIterations?: number; maxChange?: number }
  /**
   * The other workbooks that the workbook's formulas refer to (`[1]Sheet1!A1`, `'[1]My Sheet'!A1:B9`, `[1]!Rate`):
   * element `i` is `[i+1]`. References to them use the values stored here, as Excel calculates a file it opens without
   * updating its links; the other workbook itself is never read. The list is included in `inputHash` when it is not
   * empty.
   */
  externalBooks?: ExternalBook[]
  /**
   * The workbook's default font, its Normal style's (in xlsx the first font), in points: what a cell shows whose style
   * gives no font name or size. Without it, Calibri 11. It is kept, read and written, and the engine measures as for
   * Calibri 11 whatever it is: column widths in characters, fitted row heights, `####` in a column too narrow and
   * `entryWidth`, which Excel measures in the default font, so for another font they can differ from Excel's. It is
   * included in `inputHash` when given.
   */
  font?: { name: string; size: number }
}

/**
 * An external workbook, as stored alongside the formulas that refer to it (xlsx `externalLinkN.xml`): its location, and
 * the values Excel saved from it when it last read it.
 */
export interface ExternalBook {
  /** Where the link points, as stored in the file: a file name, a relative or absolute path, or a `file:` or web address. */
  target: string
  /** Its sheets in the order the file lists them. */
  sheets: ExternalSheet[]
  /** Its defined names that the workbook's formulas use: workbook-level ones, or ones local to one of its sheets. */
  names?: ExternalName[]
}

/** A sheet of an external workbook and the values stored from it. */
export interface ExternalSheet {
  name: string
  /**
   * The stored values by A1 address, null for a cell stored as empty. A cell missing from them is blank, or `#REF!` with
   * `refreshError`. Without `cells`, nothing from the sheet is stored, and every cell on it is `#REF!`.
   */
  cells?: Record<string, ExternalValue | null>
  /**
   * Excel's flag that the last update of the sheet's values failed (`refreshError`): a cell missing from them is `#REF!`
   * instead of blank. Excel sets it on every sheet of a closed workbook when a formula referring to that workbook is
   * entered, and so does a session.
   */
  refreshError?: true
}

/** A stored value of an external workbook's cell. */
export type ExternalValue = number | string | boolean | { error: ErrorCode }

/** A defined name of an external workbook. */
export interface ExternalName {
  name: string
  /** What it refers to, as stored in the file without `=`: a reference into its workbook (`Sheet1!$B$1`) or `#REF!`. */
  ref: string
  /** The index in `sheets` of the sheet the name is local to; a workbook-level name has none. */
  sheet?: number
}

/**
 * A workbook as returned by the calls that return one (`Session.getBook`, `getSnapshot`, `getLayout`, `editStructure`,
 * `readXlsx`): a `BookInput`, which can be passed back to any call that accepts a workbook.
 */
export type Book = BookInput

/** A sheet of a `Book`. */
export type Sheet = SheetInput

/** A cell of a `Book`. */
export type Cell = CellInput

/** A sheet without its cells and their formatting (`Session.getLayout`). */
export type SheetLayout = Omit<Sheet, 'cells' | 'looks'>

/** A workbook without the cells of its sheets (`Session.getLayout`): its styles, names and settings, and each sheet's layout. */
export type BookLayout = Omit<Book, 'sheets'> & { sheets: SheetLayout[] }
