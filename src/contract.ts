/**
 * The formula engine contract (subpath `./contract`): types only, no engine code. Any engine
 * implementation (currently formulas compiled to closures or to JS text) sits behind this contract; the
 * package's behaviour tests are its acceptance suite.
 *
 * A formula is stored in the `formula` field without the leading `=` (one given is dropped), separately from the value:
 * a string from data never becomes a formula. Function names are English, arguments are separated
 * by commas, the decimal separator is a dot.
 */
import type { Locale } from './locale/types.js'

export type { Locale } from './locale/types.js'

export type Scalar = number | string | boolean | null

export type JsonValue = Scalar | JsonValue[] | { [key: string]: JsonValue }

/** An input cell. A date is passed explicitly (`{ date: '2025-01-31' }`) and becomes an Excel serial number. */
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
  /** @deprecated The name is `numberFormat`. */
  format?: NumberFormat
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

/** A colour, `#RRGGBB`. */
export type Color = string

/** A font; what a style leaves out is the book's default font. */
export interface FontStyle {
  name?: string
  /** Points, 1 to 409. */
  size?: number
  bold?: boolean
  italic?: boolean
  underline?: 'single' | 'double'
  strike?: boolean
  color?: Color
}

/** One side of a cell's border; the kinds are xlsx's. */
export interface BorderLine {
  style: 'thin' | 'medium' | 'thick' | 'dotted' | 'dashed' | 'double' | 'hair' | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot'
  color?: Color
}

export interface AlignmentStyle {
  horizontal?: 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed'
  vertical?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed'
  wrap?: boolean
  /** Indent level, 0 to 250. */
  indent?: number
}

/**
 * A cell style of the book's table (`BookInput.styles`), as an xlsx cell format without its number format
 * (that is the cell's `numberFormat`). Styles do not take part in computation or in the values hash.
 */
export interface CellStyle {
  font?: FontStyle
  /** A solid fill. */
  fill?: { color: Color }
  border?: { top?: BorderLine; right?: BorderLine; bottom?: BorderLine; left?: BorderLine }
  alignment?: AlignmentStyle
}

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
  /** @deprecated The name is `numberFormat`. */
  format?: NumberFormat
  /** Hidden (xlsx `hidden`); a column of width 0 is hidden too. SUBTOTAL and AGGREGATE count the cells of hidden columns all the same, as Excel does. */
  hidden?: boolean
  /** The outline level (Excel's Group, xlsx `outlineLevel`), 1 to 7; a column without one is not grouped. */
  outlineLevel?: number
  /** The column's button collapses a group (xlsx `collapsed`): the summary column of a group whose columns are hidden. */
  collapsed?: boolean
}

/** Rows of a sheet, as xlsx `<row>`: the height, and the style and format of their cells that have none. */
export interface RowInput {
  /** Points, 0 to 409. */
  height?: number
  style?: number
  numberFormat?: NumberFormat
  /** @deprecated The name is `numberFormat`. */
  format?: NumberFormat
  /**
   * Hidden (xlsx `hidden`), by hand or by a filter: SUBTOTAL with codes 101–111 and AGGREGATE with options 1, 3, 5
   * and 7 skip the row's cells, and on a `filtered` sheet SUBTOTAL with codes 1–11 too, as Excel does. A row of
   * `height` 0 is hidden too, as Excel takes it. A collapsed group's rows are hidden rows.
   */
  hidden?: boolean
  /** The outline level (Excel's Group, xlsx `outlineLevel`), 1 to 7; a row without one is not grouped. */
  outlineLevel?: number
  /** The row's button collapses a group (xlsx `collapsed`): the summary row of a group whose rows are hidden. */
  collapsed?: boolean
}

export interface SheetInput {
  name: string
  /** Cells by A1 address (`B12`). */
  cells: Record<string, CellInput>
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
 * A conditional format: the cells it applies to, a condition and what a cell that meets it shows. The rules that hold
 * for a cell apply in the sheet's order: a field of a style an earlier rule sets wins, and `stopIfTrue` leaves out
 * the rules after one that holds. Its formulas are stored text, as a cell's: a leading `=` is dropped.
 */
export type ConditionalFormat =
  | CellIsFormat
  | ExpressionFormat
  | TextFormat
  | BlanksFormat
  | TimePeriodFormat
  | TopFormat
  | AverageFormat
  | DuplicatesFormat
  | ColorScaleFormat
  | DataBarFormat
  | IconSetFormat

export interface ConditionalBase {
  /**
   * The cells it applies to: rectangles of the sheet (`B2:B20`, `D2`, whole columns `F:F`, whole rows `3:5`), as Excel's
   * "Applies to" and xlsx `sqref`. Its formulas are written for the top-left cell of the rectangle around them: a
   * reference without `$` moves from there to each cell, as a copied formula does.
   */
  ranges: string[]
  /** Excel's Stop If True: the rules after this one do not apply to a cell this one holds for. */
  stopIfTrue?: boolean
}

/** What a rule gives the cells it holds for, as an xlsx differential format (`dxf`): each field it sets replaces the cell's own. */
export interface ConditionalStyle {
  /** `bold: false` takes away a bold the cell has, and so on. */
  font?: Pick<FontStyle, 'bold' | 'italic' | 'underline' | 'strike' | 'color'>
  fill?: { color: Color }
  border?: CellStyle['border']
  /** The number format the cell shows its value by. */
  numberFormat?: NumberFormat
  /** @deprecated The name is `numberFormat`. */
  format?: NumberFormat
}

/**
 * The cell's value compared with one formula's value, or two for `between` and `notBetween` (both ends included, either
 * way round), as `=`, `<` and the like compare; a blank cell a formula refers to is 0 there.
 */
export interface CellIsFormat extends ConditionalBase {
  type: 'cellIs'
  operator: 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual' | 'between' | 'notBetween'
  /** Formulas as a cell stores them, without `=`: `100`, `"done"`, `$B$1`. */
  formulas: string[]
  style?: ConditionalStyle
}

/** A formula that holds, as IF reads its condition: TRUE, a number other than 0, the text TRUE. */
export interface ExpressionFormat extends ConditionalBase {
  type: 'expression'
  formula: string
  style?: ConditionalStyle
}

/**
 * The cell's text holds `text`, does not, begins or ends with it, in any case. Excel computes the rule's formula, which
 * it writes beside the text (`NOT(ISERROR(SEARCH("ab",A1)))`); without `formula` the rule has that one.
 */
export interface TextFormat extends ConditionalBase {
  type: 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith'
  text: string
  formula?: string
  style?: ConditionalStyle
}

/**
 * The cell is blank (or holds spaces alone), is not, holds an error, or does not. Blanks go by the value, as Excel shows
 * them (an error is not blank); errors by `formula`, Excel's (`ISERROR(A1)`) without it, which Excel writes beside both.
 */
export interface BlanksFormat extends ConditionalBase {
  type: 'containsBlanks' | 'notContainsBlanks' | 'containsErrors' | 'notContainsErrors'
  formula?: string
  style?: ConditionalStyle
}

/** The cell holds a date in the period, by the book's clock (`TODAY()`); computed by `formula`, Excel's without it. */
export interface TimePeriodFormat extends ConditionalBase {
  type: 'timePeriod'
  period: 'today' | 'yesterday' | 'tomorrow' | 'last7Days' | 'thisWeek' | 'lastWeek' | 'nextWeek' | 'thisMonth' | 'lastMonth' | 'nextMonth'
  formula?: string
  style?: ConditionalStyle
}

/** The `rank` highest numbers of the cells (the lowest with `bottom`), or the `rank` percent of them with `percent`. */
export interface TopFormat extends ConditionalBase {
  type: 'top10'
  /** 1 to 1,000; 1 to 100 with `percent`. */
  rank: number
  percent?: boolean
  bottom?: boolean
  style?: ConditionalStyle
}

/** Numbers above the average of the cells' numbers (below it with `below`, or equal too with `equal`), or `stdDev` standard deviations past it. */
export interface AverageFormat extends ConditionalBase {
  type: 'aboveAverage'
  below?: boolean
  equal?: boolean
  /** 1 to 3. */
  stdDev?: number
  style?: ConditionalStyle
}

/** Values that occur more than once among the cells (text in any case), or once. */
export interface DuplicatesFormat extends ConditionalBase {
  type: 'duplicateValues' | 'uniqueValues'
  style?: ConditionalStyle
}

/**
 * A threshold of a colour scale, a data bar or an icon set, over the numbers of the cells the rule applies to. `min`
 * and `max`: the lowest and the highest number; `num`: the number `value` gives; `percent`: `value` percent of the
 * way from the lowest to the highest; `percentile`: the `value`-th percentile of the numbers; `formula`: what `value`
 * computes. A data bar's `autoMin` and `autoMax` are Excel 2010's automatic ends: 0, or the lowest number below 0
 * (the highest above it).
 */
export interface ConditionalValue {
  type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula' | 'autoMin' | 'autoMax'
  /** A formula as a cell stores it (`10`, `$B$1*2`), for `num`, `percent`, `percentile` and `formula`. */
  value?: string
}

/** Each number gets a colour between those of the points around it; other cells get none. */
export interface ColorScaleFormat extends ConditionalBase {
  type: 'colorScale'
  /** Two or three points, from the lowest. */
  points: { value: ConditionalValue; color: Color }[]
}

/**
 * Each number gets a bar across its cell, as long as it is far between `min` and `max`. The fields from `gradient` on
 * are Excel 2010's (xlsx `x14:dataBar`).
 */
export interface DataBarFormat extends ConditionalBase {
  type: 'dataBar'
  min: ConditionalValue
  max: ConditionalValue
  color: Color
  /** False: the cell shows the bar without its value. */
  showValue?: boolean
  /** The shortest and the longest bar in percent of the cell's width: 0 and 100 by default, 10 and 90 for an Excel 2007 bar. */
  minLength?: number
  maxLength?: number
  /** A gradient fill (true by default) or a solid one. */
  gradient?: boolean
  /** The bar's border; none without it. */
  border?: Color
  /** The bars' direction; by default the sheet's (left to right). */
  direction?: 'leftToRight' | 'rightToLeft'
  /** The fill of bars of negative numbers; the bar's colour without it. */
  negativeColor?: Color
  /** Their border; `border` without it. */
  negativeBorder?: Color
  /** Where the axis between negative and positive bars stands: by the numbers (by default), at the middle of the cell, or none (negative bars go the same way). */
  axis?: 'middle' | 'none'
  /** The axis's colour; black without it. */
  axisColor?: Color
  /** A bar of Excel 2007 (a file without Excel 2010's part of it): a gradient without a border, negative numbers as shorter bars, no axis; the fields from `gradient` on do not apply. */
  excel2007?: boolean
}

/** Excel's icon sets by their xlsx names; `3Stars`, `3Triangles` and `5Boxes` are Excel 2010's. */
export type IconSetName =
  | '3Arrows'
  | '3ArrowsGray'
  | '3Flags'
  | '3TrafficLights1'
  | '3TrafficLights2'
  | '3Signs'
  | '3Symbols'
  | '3Symbols2'
  | '3Stars'
  | '3Triangles'
  | '4Arrows'
  | '4ArrowsGray'
  | '4RedToBlack'
  | '4Rating'
  | '4TrafficLights'
  | '5Arrows'
  | '5ArrowsGray'
  | '5Rating'
  | '5Quarters'
  | '5Boxes'

/** An icon of a set: the set and its number in it, from 0, in Excel's order (xlsx `iconId`: for `3Arrows` 0 is the red arrow down, 2 the green arrow up). */
export interface Icon {
  set: IconSetName
  index: number
}

/** Each number gets an icon of the set by the thresholds it reaches. */
export interface IconSetFormat extends ConditionalBase {
  type: 'iconSet'
  icons: IconSetName
  /**
   * One threshold per icon of the set, from the lowest; a number gets the icon of the last threshold it reaches (`gte`:
   * at or above it, true by default; false: above it), the first threshold always. `icon` puts another icon there,
   * `null` none (Excel 2010's custom icons).
   */
  points: { value: ConditionalValue; gte?: boolean; icon?: Icon | null }[]
  /** The icons in the reverse order: the last for the lowest numbers. */
  reverse?: boolean
  /** False: the cell shows the icon without its value. */
  showValue?: boolean
}

/**
 * What a sheet's conditional formats give a cell (`CellView.appearance`), by the values the cells hold now. A field the
 * cell shows otherwise than by its own style and value is set.
 */
export interface ConditionalAppearance {
  /** The styles of the rules that hold, merged in their order (the first rule's field wins); a colour scale's colour is `fill`. */
  style?: ConditionalStyle
  bar?: BarLook
  icon?: Icon
  /** A data bar or an icon set hides the value: the cell shows its bar or icon alone. */
  hideValue?: true
}

/** @deprecated The name is `ConditionalAppearance`. */
export type ConditionalLook = ConditionalAppearance

/** A data bar in a cell: where it is drawn, as fractions of the cell's width from the side its direction starts at. */
export interface BarLook {
  /** From 0 to 1; `start` < `end`. */
  start: number
  end: number
  color: Color
  gradient: boolean
  border?: Color
  /** The axis between negative and positive bars, a fraction of the width, for a rule with negative numbers. */
  axis?: number
  axisColor?: Color
  /** The bars run from right to left: the fractions count from the cell's right edge. */
  rightToLeft?: true
}

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
 * What a rule of data validation lets a cell hold: any value (`any`, a rule for its input message alone), a whole number,
 * a number, a date or a time compared with the rule's formulas, a text whose length compares so, one of a list, or a value
 * for which a formula (`custom`) is TRUE.
 */
export type ValidationType = 'any' | 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | 'list' | 'custom'

/** How a value compares with `formula1` (and `formula2`): between and not between them take both, inclusive. */
export type ValidationOperator = 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual'

/**
 * A rule of data validation, as xlsx `<dataValidation>`. Its formulas are stored formula text without `=` (en-US, as a
 * cell's; a leading `=` is dropped), relative to the rule's first cell, the top-left cell of the first range of `range`:
 * a rule over `B2:B9` with `formula1: 'A2>0'` checks `A5>0` for B5. A literal list is a text constant of items separated by commas
 * (`'"Yes,No,Maybe"'`), as a file stores it in every locale; a list from cells is a reference (`'$D$2:$D$9'`,
 * `"'Lists'!$A$1:$A$20"`), a defined name or a formula giving one.
 */
export interface DataValidation {
  /** The cells: ranges separated by spaces (`B2:B20 D2`, `C:C`), as xlsx `sqref`; no other rule of the sheet meets them. */
  range: string
  type: ValidationType
  /** For whole, decimal, date, time and textLength: how the value compares; `between` by default. */
  operator?: ValidationOperator
  /** The value, the low bound, the list or the formula of a custom rule; every type but `any` has one. */
  formula1?: string
  /** The high bound of `between` and `notBetween`. */
  formula2?: string
  /**
   * A blank cell passes, and so does any value while a bound or a list refers to a blank cell (Excel's Ignore blank,
   * xlsx `allowBlank`); true by default, as Excel makes a rule.
   */
  allowBlank?: boolean
  /** A list shows its dropdown in the cell; true by default (false is xlsx `showDropDown="1"`). */
  dropdown?: boolean
  /** The input message Excel shows while the cell is selected: its title (Excel takes 32 characters) and text (255). */
  promptTitle?: string
  prompt?: string
  /** The input message is shown; true by default, as Excel makes a rule. */
  showPrompt?: boolean
  /** The alert for a value the rule does not take: stop (the default), warning or information, and its title (32) and text (225). */
  errorStyle?: 'stop' | 'warning' | 'information'
  errorTitle?: string
  error?: string
  /** The alert is shown; true by default, as Excel makes a rule. Without it Excel takes any value typed. */
  showError?: boolean
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

export interface BookInput {
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
   * Iterative calculation of circular references, as Excel's setting and xlsx `calcPr iterate`. Without it a cycle
   * gives `#CYCLE!`. With it the cells of every cycle and the cells that depend on them are computed from 0, again
   * and again, sheet by sheet in the order of the sheets' names (in any case), row by row. That stops after
   * `maxIterations` rounds (100 by default, 1 to 32,767) or when no value changes by `maxChange` or more (0.001 by
   * default). Each computation starts from 0 again, so a session gives what a full computation gives. Excel goes on
   * from the values it holds instead.
   */
  iteration?: { maxIterations?: number; maxChange?: number }
}

/**
 * `CALC`: Excel's #CALC!, an empty array (`FILTER` that keeps nothing) or a text over 32,767 characters (`CONCAT`).
 * `SPILL`: an array that cannot spill from its formula's cell (its range is not empty, meets a merged area or
 * leaves the sheet). `NULL`: Excel's #NULL!, ranges that do not meet under the intersection operator (`A1:A3 C1:C3`).
 * `ERROR`: a formula that does not compute (it does not parse, say).
 */
export type ErrorCode = 'DIV0' | 'NA' | 'NAME' | 'NUM' | 'REF' | 'VALUE' | 'NULL' | 'CYCLE' | 'SPILL' | 'CALC' | 'ERROR'

/** A computed value. Dates come out as serial numbers, numbers are rounded to 15 significant digits. */
export type CellValue = Scalar | { error: ErrorCode }

/**
 * Formula check codes. `unsupported`: an external function called with arguments that are not literals.
 * `unknown_name` covers tables too: an unknown table or column, a structured reference without a table outside one.
 */
export type CheckCode =
  | 'empty'
  | 'syntax'
  | 'too_long'
  | 'too_deep'
  | 'external_ref'
  | 'unknown_function'
  | 'unknown_sheet'
  | 'unknown_name'
  | 'unsupported'

/** The moment (ISO) and the random seed that TODAY, NOW, RAND and RANDBETWEEN are computed from. */
export interface Clock {
  now: string
  seed: number
}

export interface CellRef {
  sheet: string
  cell: string
}

/**
 * External functions (host functions evaluated outside the book): their values are resolved in
 * advance, in one request, and passed to the computation. The key of a call is its canonical text
 * (`BALANCE("2025-01")`); arguments are literals only.
 */
export interface Externals {
  functions: string[]
  values?: Record<string, CellValue>
}

/**
 * How formulas run; values, issues and hashes are the same either way. `closures` (the default) builds each
 * formula as a tree of JavaScript closures and generates no code, so it works under a Content Security
 * Policy without 'unsafe-eval'. `eval` generates JavaScript text and compiles it with `new Function`, which
 * such a policy blocks (computing then throws an `EvalError`).
 */
export type Execution = 'closures' | 'eval'

/** A value as a cell shows it: its text and the colour of its format's section. */
export interface CellDisplay {
  /**
   * What Excel shows in the cell (VBA's `Range.Text`): the value by its number format in the book's locale, General
   * narrowed to the column's width, `#` across the column for a number that does not fit, fills `*x` repeated
   * across it; text as it is (it overflows into the next cells). Derived from the value, the format, the locale and
   * the width: not part of the values hash.
   */
  text: string
  /** The colour of the format's section (`[Red]`, `[Color10]`), `#RRGGBB`; absent without one. */
  color?: Color
}

/** `displayValue`'s options: where the value is shown. */
export interface DisplayOptions {
  /** The number format; General without it. */
  numberFormat?: NumberFormat
  /** @deprecated The name is `numberFormat`. */
  format?: NumberFormat
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
   * The cell typed into, as the book holds it (a `CellInput`, or a `CellView` of `Session.getCells`): its `numberFormat` (the one
   * it shows by: its row's or column's without its own), `style`, `meta`, `note` and `link`; other fields are not read.
   * The new input keeps its style, meta, note and link, and its format where Excel keeps it.
   */
  cell?: { numberFormat?: NumberFormat; /** @deprecated The name is `numberFormat`. */ format?: NumberFormat; style?: number; meta?: JsonValue; note?: CellNote; link?: CellLink } | null
  /** The moment whose year completes a date typed without one (`15.3` in ru-RU, `3/15` in en-US), ISO as `Clock.now`; the current moment without it. */
  now?: string
}

/**
 * Text typed into a cell as Excel takes it (`parseEntry`): the cell's new input, `null` when nothing is left of the
 * cell; or a formula Excel does not take, with a check code (`syntax`, `too_long`, `too_deep`), a message in the
 * language of `opts.locale` (English for en-US) and the position in the typed text.
 */
export type EntryResult = { ok: true; input: CellInput | null } | (Refusal<CheckCode> & { at: number })

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

/** A cell's row and column, numbered from 1 as A1 numbers them: `B12` is row 12, column 2. */
export interface CellPosition {
  row: number
  col: number
}

/** A rectangle of a sheet: rows and columns numbered from 1, bounds inclusive; a whole column runs from row 1 to `MAX_ROWS`. */
export interface CellRect {
  top: number
  left: number
  bottom: number
  right: number
}

/** A range's rectangle and what its text names: cells (`B2:D9`, `B2`), whole columns (`B:D`) or whole rows (`2:9`). */
export interface RangeRect extends CellRect {
  kind: 'cells' | 'columns' | 'rows'
}

/** The lines a key of `SheetInput.columns` or `rows` names (`C`, `B:D`, `5`, `5:9`), numbered from 1. */
export interface LineSpan {
  first: number
  last: number
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

/** `scanReferences`'s options. */
export interface ScanOptions {
  /** The locale the text is typed in, a loaded locale's tag (`LOCALES`) or a locale object; without it the text is a stored formula. */
  locale?: string | Locale
}

/**
 * Size limits of a computation. A book over them is not computed: the call throws `CalcError` with code `limit`, so a
 * result is either exact or absent and the values hash never depends on the limits.
 */
export interface CalcLimits {
  /** Formulas in the book. */
  formulas?: number
  /** Values of one range read whole or one array that a formula makes, checked before they are made. */
  cells?: number
  /**
   * Values one spill may fill (its stored cells: the cells past a large array's last distinct value are not stored),
   * checked before they are written; `cells` when not given.
   */
  spill?: number
}

export interface ComputeOptions {
  /**
   * The clock for TODAY, NOW, RAND and RANDBETWEEN. By default the current moment and a random seed;
   * pass the `clock` of a result to reproduce that result exactly.
   */
  clock?: Clock
  externals?: Externals
  /** `closures` by default. */
  execution?: Execution
  /** Each cell's shown text and colour (`CellResult.text`, `color`); true by default, false computes values alone. */
  display?: boolean
  /** None by default. */
  limits?: CalcLimits
}

/** A cell's value, and by default (`display`) what the cell shows: `text` and `color`, as `CellDisplay`. */
export interface CellResult extends Partial<CellDisplay> {
  /** The value: numbers stay numbers whatever the format shows. */
  value: CellValue
  /**
   * Canonical formula text (sheet always quoted, function names upper-case), with sheets, tables and defined names
   * spelled as the book names them, as Excel shows them (`ставка*2` with the name `Ставка` is `Ставка*2`).
   */
  formula?: string
  /** The cell's own number format, as its input has it. */
  numberFormat?: NumberFormat
  meta?: JsonValue
  /**
   * A formula whose array spilled: the range it fills on its sheet (`B2:D9`, this cell at the top-left); for an array
   * formula of the old kind (`CellInput.array`), its area.
   */
  spill?: string
  /**
   * The value of the cells of `spill` that the result does not list: only for an array of more than 65,536
   * values whose cells past some row and column all hold one value (`A:A*2`: 0 below the data), which are
   * then left out of `cells` and of the hash's lines but for one line of the range.
   */
  spillFill?: CellValue
  /**
   * A cell an array spilled into, or a cell of an old-kind array formula's area: the address of the array's formula
   * on the same sheet (an empty element is 0).
   */
  spillOf?: string
}

/** A cell with an error and its reason: an invalid formula (with the check code and position) or a computation error. */
export interface CellIssue extends CellRef {
  error: ErrorCode
  reason: string
  check?: CheckCode
  /** Position in the formula text for `check`. */
  at?: number
}

export interface NameIssue {
  name: string
  scope?: string
  reason: string
  /**
   * The book has no such name: one Excel would not take, a duplicate, an unknown scope sheet, a formula that does not
   * parse or a reference without its sheet; a formula using the word finds no name. False: the name is in the book and
   * gives an error where it is used (an unknown sheet, table or function in its formula).
   */
  dropped: boolean
}

/**
 * What a result reports about its book, by `kind`: a cell with an error and its reason (`cell`, a `CellIssue`), or a
 * name the book left out or keeps with an error (`name`, a `NameIssue`).
 */
export type Diagnostic = ({ kind: 'cell' } & CellIssue) | ({ kind: 'name' } & NameIssue)

export interface ComputeResult {
  sheets: { name: string; cells: Record<string, CellResult> }[]
  /** Every cell with an error, sheet by sheet in the book's order and row by row, then every name with a problem, in the order of the book's names. */
  diagnostics: Diagnostic[]
  /** Cells that depend on volatile functions (directly or through other cells). */
  volatile: CellRef[]
  /** The clock the book was computed with (for a live session, the latest one). */
  clock: Clock
  /** sha256 of the canonical text of the values (spilled cells too) together with the engine configuration version. */
  hash: string
  /** sha256 of the input (sheets, cells, names, and styles, columns, rows, merged areas, frozen panes, conditional formats, data validation and outline settings when there are any; order does not matter). */
  inputHash: string
  configVersion: string
  ms: number
}

export interface ValidateContext {
  /** Sheets of the book; without them sheets are not checked. */
  sheets?: string[]
  /** Names of the book; without them names are not checked. */
  names?: string[]
  externalFunctions?: string[]
  /**
   * Tables of the book with their columns; without them structured references are not checked. With them the canonical
   * text spells tables and columns as they are named, and a bare table name is a table, not a name.
   */
  tables?: { name: string; columns: string[] }[]
}

export type ValidateResult =
  /** `tables`: the tables the formula's structured references name (and bare table names, with `ctx.tables`). */
  | { ok: true; canonical: string; functions: string[]; sheets: string[]; names: string[]; externals: string[]; tables: string[] }
  | (Refusal<CheckCode> & { at: number })

/**
 * A check or an edit that did not pass: a code a program reads, a message for people (in the book's locale where the
 * call has one, English otherwise) and, for typed text, the position in it. Every check answers `{ ok: true, … }` or
 * a `Refusal` with its own codes.
 */
export interface Refusal<Code extends string = string> {
  ok: false
  code: Code
  message: string
  at?: number
}

/**
 * Why Excel would not take a name (`validateName`). `chars`: it does not start with a letter of any script, `_` or `\`,
 * holds a character other than letters, digits, `_`, `.`, `\` and `?`, or is empty or longer than 255 characters.
 * `reference`: it reads as a cell address or an R1C1 reference (`Q1`, `R`, `R2C3`). `boolean`: TRUE or FALSE, or the
 * book's locale's word for one, in any case (`ИСТИНА` in ru-RU).
 */
export type NameRefusal = 'chars' | 'reference' | 'boolean'

export type NameCheck = { ok: true } | Refusal<NameRefusal>

/** The book whose rules `validateName` checks a name by. */
export interface NameOptions {
  /** The book's locale (`BookInput.locale`): a loaded locale's tag, or a locale object; en-US without one. */
  locale?: string | Locale
}

/** The parts of a book `validateNames` checks names against. */
export interface NamesContext {
  /** Sheets of the book: a name's scope must be one of them, and its formula's references to others give `#REF!`. */
  sheets: string[]
  /** Tables of the book with their columns, for the structured references of names' formulas. */
  tables?: { name: string; columns: string[] }[]
  externalFunctions?: string[]
  /** The book's locale (`BookInput.locale`): a loaded locale's tag, or a locale object; en-US without one. */
  locale?: string | Locale
}

/** Fields of a cell to set (`Session.setCells`): a value sets the field, null removes it, a field left out stays. */
export type CellFields = { [K in keyof CellInput]?: CellInput[K] | null }

/**
 * A cell edit (`Session.setCells`): the fields given set the cell's, null removes one, and the fields left out stay, as
 * Excel keeps a cell's look when a value is typed over it. The content is one: a new `value` removes the formula and its
 * `array`, a new `formula` the value. `input: null` clears the content and the hyperlink and keeps the rest, as Excel's
 * Delete key.
 */
export interface CellEdit extends CellRef {
  input: CellFields | null
}

/** A cell replaced whole (`Session.replaceCells`): its new input, or null to remove the cell. */
export interface CellReplace extends CellRef {
  input: CellInput | null
}

/**
 * A changed cell, with `spill`, `spillFill` and `spillOf`, and with `display` its `text` and `color`, as in
 * `CellResult`. A formula whose spill range or `spillFill` changed is listed even when its own value did not; the
 * cells it no longer lists are listed when their value changed. With `display`, a cell whose shown text or colour
 * changed is listed too (a new format, a column's width), with its value.
 */
/**
 * A cell of `Session.getCells`: as `CellResult`, and the cell's own `style` (an index of the book's `styles`), `note` and
 * `link`, and the `appearance` its sheet's conditional formats give it (with a number format of theirs, `text` and
 * `color` show the value by it).
 */
export interface CellView extends CellResult {
  style?: number
  note?: CellNote
  link?: CellLink
  appearance?: ConditionalAppearance
  /** On the top-left cell of an array formula of the old kind, its area (`CellInput.array`); `spill` is that area too. */
  array?: string
}

export interface CellChange extends CellRef, Partial<CellDisplay> {
  value: CellValue
  spill?: string
  spillFill?: CellValue
  spillOf?: string
  /** The cell's formula, when it changed: canonical text, as `CellResult.formula`; null when the cell has none now. */
  formula?: string | null
  /** The cell's own style, when it changed, as `CellView.style`; null when it has none now (it shows its row's or column's). */
  style?: number | null
  /** The cell's own number format, when it changed; null when it has none now. */
  numberFormat?: NumberFormat | null
}

/**
 * What a call that changes a session's book answers: the session's revision after it, the cells it changed, and what
 * else a host showing the book reads again.
 */
export interface EditResult {
  /** The session's revision after the call: every call that changes the book adds one, `undo` and `redo` too. */
  revision: number
  /**
   * The cells whose value, shown text, formula or look changed, by `CellChange`; a cell another cell's value changed
   * is listed with its value alone. After a structure operation, the cells whose value or formula differ from those of
   * the cell that was at their old place (`moved`).
   */
  changes: CellChange[]
  /** The sheets whose layout changed (rows, columns, merged areas, tables, conditional formats, validation…): a host reads `getLayout()` again. */
  layout?: string[]
  /** The book's style table, when the call changed it. */
  styles?: CellStyle[]
  /** How addresses moved, in order, after a structure operation or its undo: a host moves its window by them. */
  moved?: AddressMove[]
  /**
   * Rectangles where cells may look otherwise without being in `changes`: the cells of the book's conditional formats,
   * after a call that changed a value or the rules (a rule's look reads values anywhere its formulas point).
   */
  repaint?: RangeRef[]
}

/**
 * How a structure operation moved a book's addresses (`EditResult.moved`). `lines`: the lines from `at` (a row's number,
 * or a column's, A being 1) on move by `count`: down or right by the lines inserted before `at`, up or left when
 * `-count` lines from `at` on were deleted. `range`: the cells of `from` moved so that its top-left cell is at `to`,
 * what was there gone. A sheet renamed, deleted, inserted or moved: `index`, its place among the sheets, from 0.
 */
export type AddressMove =
  | { kind: 'lines'; sheet: string; axis: 'rows' | 'columns'; at: number; count: number }
  | { kind: 'range'; from: RangeRef; to: CellRef }
  | { kind: 'renameSheet'; sheet: string; name: string }
  | { kind: 'deleteSheet'; sheet: string }
  | { kind: 'insertSheet'; sheet: string; index: number }
  | { kind: 'moveSheet'; sheet: string; index: number }

/** Options of a call that changes a session's book. */
export interface EditOptions {
  /**
   * The revision the caller's view of the book is of (`EditResult.revision`, `Session.revision`): when the session's is
   * another, the call changes nothing and throws `CalcError` with code `stale_revision`, so that a host does not apply
   * an edit over one it has not seen.
   */
  baseRevision?: number
}

/** What `Session.clearCells` clears: the content and hyperlinks (Excel's ClearContents), the look (ClearFormats), or both and the rest (Clear). */
export type ClearKind = 'contents' | 'formats' | 'all'

/** `Session.setCellText`'s options. */
export interface CellTextOptions extends EditOptions {
  /** False: the width a typed number widens its column to is given (`width`), not set; true by default, as Excel widens it. */
  widen?: boolean
}

/** Why `Session.setCellText` did not take the text: a formula Excel does not take (`CheckCode`), part of an array formula, a formula in a table's header cell. */
export type CellTextRefusal = 'syntax' | 'empty' | 'too_long' | 'too_deep' | 'array' | 'header'

/** A value a cell's rule of data validation does not take (`Session.setCellText`): the rule, its alert's style and message. */
export interface ValidationAlert {
  valid: false
  rule: DataValidation
  severity: 'stop' | 'warning' | 'information'
  message: string
}

/**
 * Text typed into a cell (`Session.setCellText`): the edit's result, with the width the column took (`width`) and the
 * rule of data validation the value breaks (`validation`); or why the text was not taken, with its position in the text.
 */
export type CellTextResult = (EditResult & { ok: true; width?: number; validation?: ValidationAlert }) | Refusal<CellTextRefusal>

/** What `Session.setCellText` would do (`Session.previewCellText`): the cell's input, its value, the width and the alert, with nothing changed. */
export type CellTextPreview = { ok: true; input: CellInput | null; value: CellValue; width?: number; validation?: ValidationAlert } | Refusal<CellTextRefusal>

/** `Session.getViewport`'s options. */
export interface ViewportOptions {
  /** Each formula also as the book's locale shows it (`ViewportCell.formulaLocal`); false by default. */
  formulaLocal?: boolean
  /** Each cell with an error also with its reason and the cell it began in (`ViewportCell.issue`); false by default. */
  diagnostics?: boolean
  /** The most cells a window gives, 20,000 by default: of a larger one, its first rows that fit (`Viewport.truncated`). */
  maxCells?: number
}

/**
 * A cell of a window (`Session.getViewport`): its value as `getResult()` gives it, with the cell's whole look resolved, its
 * style and number format as rows and columns give them, what its conditional formats show, its note and link.
 */
export interface ViewportCell extends Partial<CellDisplay> {
  value: CellValue
  /** The canonical formula text, as `CellResult.formula`. */
  formula?: string
  /** With `formulaLocal`: the formula as Excel shows it in the book's locale (`Range.FormulaLocal` without `=`), inside its table as Excel shows it there (`[@Qty]`). */
  formulaLocal?: string
  spill?: string
  spillFill?: CellValue
  spillOf?: string
  /** On the top-left cell of an array formula of the old kind, its area (`CellInput.array`). */
  array?: string
  meta?: JsonValue
  /** The cell's own style, an index of the book's `styles` as `CellView.style`; its object is in `Viewport.styles`. */
  style?: number
  /**
   * The style the cell shows, an index of the book's `styles` (its object in `Viewport.styles`): its own, else its row's,
   * else its column's. It is always one of the table's: a cell where a styled row crosses a styled column shows the row's,
   * as Excel shows a file's empty cell there, and the two merged are a style of the cell's own (`Session.patchCells`).
   */
  effectiveStyle?: number
  /** The number format the cell shows by, as a code: its own, else its row's, else its column's, a built-in format by the code of the book's locale (14 is `dd.mm.yyyy` in ru-RU). */
  numberFormat?: string
  /** Where `numberFormat` comes from. */
  numberFormatSource?: 'cell' | 'row' | 'column'
  /** What the sheet's conditional formats give the cell by the current values; with a number format of theirs, `text` and `color` show the value by it. */
  appearance?: ConditionalAppearance
  note?: CellNote
  link?: CellLink
  /** With `diagnostics`, on a cell with an error: as `Session.getCellIssue`. */
  issue?: CellIssueDetail
}

/** A column of a window: its width as `ColumnInput.width`, Excel's default where it has none. */
export interface ViewportColumn {
  width: number
  hidden?: true
}

/** A row of a window: its height in points, Excel's default (15) where it has none. */
export interface ViewportRow {
  height: number
  hidden?: true
}

/** A table that meets a window (`Viewport.tables`): its name, its range and which of its rows are a header and totals. */
export interface ViewportTable {
  name: string
  /** The whole table, its header and totals rows included (`TableInput.range`). */
  range: string
  /** Its first row is a header row (`TableInput.headers`, true by default). */
  headers: boolean
  /** Its last row is a totals row (`TableInput.totals`, false by default). */
  totals: boolean
}

/** What a grid needs to draw a window of a sheet (`Session.getViewport`). */
export interface Viewport {
  /** The rectangle given (`B2:D9`): the one asked for, or its first rows when it holds more than `maxCells` cells. */
  range: string
  /** The window asked for held more than `maxCells` cells: `range` is the part given. */
  truncated: boolean
  /** By address, row by row: every cell with something to show, an empty one with its row's or column's style, an appearance of a conditional format, a note or a link too. */
  cells: Record<string, ViewportCell>
  /** The styles the window's cells have or show (`ViewportCell.style`, `effectiveStyle`), by their index in the book's `styles`: copies, the answer's own. */
  styles: Record<number, CellStyle>
  /** The columns of `range`, left to right. */
  columns: ViewportColumn[]
  /** The rows of `range`, top to bottom. */
  rows: ViewportRow[]
  /** The merged areas that meet `range`, whole. */
  merges: string[]
  /** The tables that meet `range`, whole, in the sheet's order. */
  tables: ViewportTable[]
  /** The sheet's frozen panes (`SheetInput.frozen`): its first `rows` rows and `columns` columns stay in view, 0 for none. */
  frozen: { rows: number; columns: number }
  /** Where the sheet's cells end, as `Session.getUsedRange`. */
  usedRange: SheetExtent
  /** The session's revision the window is of (`Session.revision`): a host takes the changes of later edits over it (`EditResult.revision`). */
  revision: number
}

/**
 * Why a cell holds an error (`Session.getCellIssue`): its code, reason, check code and position as the result's `diagnostics`
 * gives them, and `origin`, the cell where the error began, following the cells its reason names (`from 'S'!A2`) to one
 * whose error is its own; in a cycle, the first cell of the cycle the way reaches.
 */
export interface CellIssueDetail {
  error: ErrorCode
  reason: string
  check?: CheckCode
  at?: number
  origin: CellRef
}

export interface ExplainOptions {
  /** Depth over precedents, 3 by default. */
  depth?: number
  /** Node limit, 200 by default. */
  maxNodes?: number
  /** Expand ranges into their non-empty cells (by default a range is a single node). */
  expand?: boolean
  /** How many cells past `depth` the way from an error to the cell it began in goes on, 64 by default. */
  errorDepth?: number
}

export interface ExplainNode {
  /** `'Sheet'!B3` or `'Sheet'!D:D`. */
  ref: string
  value?: CellValue
  formula?: string
  /** A cell an array spilled into: its formula's cell (`'Sheet'!B2`), which is the node's child. */
  spillOf?: string
  /** For a range: the number of non-empty cells. */
  cells?: number
  /**
   * A cell with an error: the cell the error began in (`'Sheet'!A2`, the node's own `ref` when it began there), as
   * `CellIssueDetail.origin`. The way there is in the tree past `depth` too, and through a range by the cell it came from.
   */
  origin?: string
  children: ExplainNode[]
}

export interface Explanation {
  root: ExplainNode
  nodes: number
  truncated: boolean
  /**
   * The tree as indented text (for logs, an LLM agent or a person) in the book's locale: in ru-RU its words, numbers,
   * formulas, booleans and errors are Russian (`'План'!B2 = 480  ← ОКРУГЛ(B1*1,2;0)`); the nodes keep the stored text.
   * An error is written as the book's cells show it (`#DIV/0!`), with the cell it began in (`(error from 'S'!A2)`).
   */
  text: string
}

/**
 * Session options. A session keeps one clock, `clock` or by default the moment of opening and a random
 * seed, until `setClock`. A live session (`live: true`) takes a fresh clock and a new seed on every
 * recomputation instead. Edits are compiled in the `execution` the session was opened with (`closures` by
 * default). With `display` (the default) results and changes carry what the cells show. `limits` bound every call;
 * a call over them throws `CalcError` and leaves the session as it was. `undo` bounds the steps `Session.undo` keeps.
 */
export type SessionOptions = ({ clock?: Clock; live?: false } | { live: true }) & { externals?: Externals; execution?: Execution; display?: boolean; limits?: CalcLimits; undo?: UndoOptions }

/**
 * The steps a session keeps for `undo`: `depth` of them, 100 by default as in Excel (0 keeps none), holding `maxMb`
 * megabytes at most by the engine's estimate, 64 by default, in a pool with `maxHeapMb` its share of a thread's heap
 * (`maxHeapMb / (8 × maxSessionsPerThread)`, `docs/async.md`). Past either the oldest steps go; a step larger than
 * `maxMb` alone is not kept, and the steps before it go too, as Excel forgets what it cannot undo.
 */
export interface UndoOptions {
  depth?: number
  maxMb?: number
}

/** The calls a session can undo: those that change its book. `setClock` and a live session's new clock are not steps. */
export type UndoCall = 'setCells' | 'replaceCells' | 'clearCells' | 'patchCells' | 'setCellText' | 'setNames' | 'setStyles' | 'setLayout' | 'editStructure'

/** A step of a session undone or done again (`Session.undo`, `Session.redo`): the edit's result, as the call gave it. */
export interface UndoResult extends EditResult {
  /** The call whose change it was. */
  call: UndoCall
  /** The calls that change cells: the cells whose input changed back, so a host can read their note and link again. */
  cells?: CellRef[]
  /**
   * `editStructure`: where a cell of the book before this call is now, null when no cell of the book has it; an interface
   * that moves its shown values by it and applies `changes` has the new values, as after `editStructure`.
   */
  mapCell?(ref: CellRef): CellRef | null
}

/** What a session can undo and redo now: the steps kept each way. */
export interface SessionHistory {
  undo: number
  redo: number
}

/** A rectangle on a sheet: `B2`, `B2:D9`, whole columns `B:D` or whole rows `2:9` (no `$`, no sheet). */
export interface RangeRef {
  sheet: string
  range: string
}

/**
 * One structure operation (`editStructure`). Rows are 1-based numbers and columns are letters, as in A1
 * addresses; `count` is 1 by default; sheets and tables are matched case-insensitively. Tables move with their cells
 * and grow and shrink with the rows and columns inserted and deleted inside them; an inserted row gets the formulas of
 * the calculated columns, an inserted column a name (`Column1`), and references to a deleted column become #REF!.
 */
export type StructureOp =
  /** Inserts `count` rows before `row`: rows from `row` on move down; the new rows take the format of the row above. */
  | { op: 'insertRows'; sheet: string; row: number; count?: number }
  /** Deletes rows `row` … `row + count - 1`: the rows below move up. */
  | { op: 'deleteRows'; sheet: string; row: number; count?: number }
  /** Inserts `count` columns before `column`: columns from `column` on move right; the new ones take the format of the column on the left. */
  | { op: 'insertColumns'; sheet: string; column: string; count?: number }
  /** Deletes `count` columns from `column` on: the columns to the right move left. */
  | { op: 'deleteColumns'; sheet: string; column: string; count?: number }
  /** Cut and paste: the range moves so that its top-left cell lands on `to`, replacing what was there; cells keep their `meta`. */
  | { op: 'moveRange'; from: RangeRef; to: CellRef }
  /**
   * Copy and paste: the range is pasted at `to`; a target that is a whole multiple of the source's size is
   * tiled with copies (fill down and fill right are copies of the first row or column). Copies have no
   * `meta` unless `withMeta`.
   */
  | { op: 'copyRange'; from: RangeRef; to: RangeRef; withMeta?: boolean }
  /** A new name by Excel's rules (see `BookInput.sheets`) that no other sheet has. */
  | { op: 'renameSheet'; sheet: string; name: string }
  | { op: 'deleteSheet'; sheet: string }
  /**
   * A new empty sheet named `name` (by Excel's rules, not taken), before the sheet `before`, else after the last one.
   * A reference across sheets takes in a sheet inserted between its ends.
   */
  | { op: 'insertSheet'; name: string; before?: string }
  /**
   * The sheet moved before the sheet `before`, else after the last one; before a hidden sheet it goes before the next
   * visible one, as Excel moves it. A reference across sheets covers the sheets between its ends in the new order; an end
   * that moves stays its end until it passes the other end, then the next sheet inward takes its place.
   */
  | { op: 'moveSheet'; sheet: string; before?: string }
  /** The sheet hidden (Excel's Hide Sheet), very hidden with `very` (see `SheetInput.hidden`); a book keeps one visible sheet. */
  | { op: 'hideSheet'; sheet: string; very?: boolean }
  /** The sheet shown again (Excel's Unhide Sheet). */
  | { op: 'unhideSheet'; sheet: string }
  /** A table over cells of `sheet` (Excel's Format as Table): with a header row its header cells take the columns' names. */
  | { op: 'addTable'; sheet: string; table: TableInput }
  /** A new name for the table, by the rules of `TableInput.name`; every structured reference to it follows. */
  | { op: 'renameTable'; table: string; name: string }
  /** The table and its cells deleted (Excel's Delete Table): structured references to it become #REF!. */
  | { op: 'deleteTable'; table: string }
  /**
   * The table turned into plain cells (Excel's Convert to Range): structured references to it become references to its
   * cells with `$` and the sheet (`'S'!$B$2:$B$9`; a reference to the formula's own row keeps its row relative, `'S'!$B2`).
   */
  | { op: 'convertTable'; table: string }

/** Why Excel would refuse an operation. */
export type StructureRefusal =
  /** Cells would leave the sheet (past XFD or row 1,048,576), or a paste does not fit. */
  | 'off_sheet'
  /** A new sheet name that breaks Excel's rules or is taken (sheet names compare case-insensitively). */
  | 'sheet_name'
  /** A book keeps one visible sheet at least: its last visible sheet cannot be deleted or hidden. */
  | 'last_sheet'
  /** A very hidden sheet (`SheetInput.hidden` 'very') cannot be deleted or moved, as Excel refuses: it is shown again first. */
  | 'very_hidden'
  /** A move or copy would cut through a merged area, at its source or where it lands. */
  | 'merged'
  /** A copy that leaves out rows a filter hid, pasted over hidden rows of a filtered sheet (several areas onto several areas). */
  | 'filtered'
  /** The operation would change part of the area of an array formula of the old kind (`CellInput.array`). */
  | 'array'
  /**
   * The operation would shift or cut part of a table as Excel refuses it: rows or columns inserted into or deleted from
   * two tables at once, a table's header row deleted without the rest, a move that cuts a table, a table over another.
   */
  | 'table'
  /** A table name that breaks the rules of `TableInput.name` or that a table or a defined name of the book has. */
  | 'table_name'

/** `editStructure`'s answer. */
export type StructureResult =
  | {
      ok: true
      /** The new book; the input book is not changed, unchanged sheets and cells are shared. */
      book: BookInput
      /** Where a cell of the old book is now; null if it was deleted or overwritten. */
      mapCell(ref: CellRef): CellRef | null
    }
  /** A refusal: its code, and a message in the language of the book's locale (English for en-US). */
  | Refusal<StructureRefusal>

/** @deprecated The name is `StructureResult`. */
export type RestructureResult = StructureResult

/**
 * `Session.editStructure`: the operation's result and the edit's. `changes` are the cells whose value or formula differs
 * from those of the cell that was at their old place (`mapCell`), by new address, a cell no old cell came to counting as
 * empty before; with `display`, the cells whose shown text or colour differs too. An interface that moves its shown
 * cells by `moved` (or `mapCell`) and applies these has the new ones.
 */
export type SessionStructureResult =
  | (EditResult & {
      ok: true
      /** The session's book after the operation (as `getBook()` returns it). */
      book: BookInput
      mapCell(ref: CellRef): CellRef | null
    })
  | Refusal<StructureRefusal>

/** @deprecated The name is `SessionStructureResult`. */
export type SessionRestructureResult = SessionStructureResult

/** Fields to change on columns or rows: a value sets the field, null removes it, a field left out stays. */
export type LinePatch<T> = { [K in keyof T]?: T[K] | null }

/** A change of a sheet's layout. */
export type LayoutEdit =
  /** Columns `C` or `B:D`; the patch applies to each of them. */
  | { sheet: string; columns: string; set: LinePatch<ColumnInput> }
  | ColumnsFit
  /** Rows `5` or `5:9`; the patch applies to each of them. */
  | { sheet: string; rows: string; set: LinePatch<RowInput> }
  /** All the merged areas of the sheet. */
  | { sheet: string; merges: string[] }
  /** The sheet's filter mode (`SheetInput.filtered`). */
  | { sheet: string; filtered: boolean }
  /** The sheet's frozen panes (`SheetInput.frozen`); null unfreezes them. */
  | { sheet: string; frozen: FrozenPanes | null }
  /** All the conditional formats of the sheet (`SheetInput.conditionalFormats`), in priority order. */
  | { sheet: string; conditionalFormats: ConditionalFormat[] }
  /** All the rules of data validation of the sheet (`SheetInput.validations`). */
  | { sheet: string; validations: DataValidation[] }
  /** Where the sheet's summary rows and columns are (`SheetInput.outline`); null for Excel's default, below and right. */
  | { sheet: string; outline: OutlineSettings | null }

/**
 * Excel's AutoFit of columns `C` or `B:D` (a double click on a column's border): each column with something to show gets
 * the width of its widest cell as Excel for Mac fits it (`fitColumns`), marked `bestFit` and shown if it was hidden; a
 * column with nothing to show keeps its width. The widths are computed from the values and layout the edits before it
 * leave (`docs/formats.md`, "Fitting columns").
 */
export interface ColumnsFit {
  sheet: string
  columns: string
  fit: true
}

/**
 * A cell's data validation (`Session.getValidation`): the rule over it, the rule's formulas as at the cell (relative
 * references moved from the rule's first cell), and for a list the items its dropdown offers.
 */
export interface CellValidation {
  rule: DataValidation
  formula1?: string
  formula2?: string
  /**
   * A list's items in order, computed in the book as it is now: a literal list's items as written (`text`, and `value` the
   * same text), or the cells of the reference its formula gives, each with its value and the text it shows by its number
   * format (blank cells past the last value left out). An error where the formula gives no reference or list.
   */
  list?: { value: CellValue; text: string }[] | { error: ErrorCode }
}

/**
 * Whether a value meets a cell's rule (`Session.validateEntry`); when it does not, the rule, whose alert Excel shows,
 * its style (`errorStyle`, stop by default) and its message: the rule's `error`, else the library's English text.
 */
export type EntryCheck = { ok: true } | (Refusal<'validation'> & { rule: DataValidation; severity: 'stop' | 'warning' | 'information' })

/** Fields of a style to change (`CellPatch.set.style`): a value sets a field, null removes it, a field left out stays. */
export interface StylePatch {
  font?: { [K in keyof FontStyle]?: FontStyle[K] | null } | null
  fill?: { color: Color } | null
  border?: { [K in keyof NonNullable<CellStyle['border']>]?: BorderLine | null } | null
  alignment?: { [K in keyof AlignmentStyle]?: AlignmentStyle[K] | null } | null
}

/**
 * A change of the look of a cell (`cell`) or a range (`range`: `B2:D9`, whole columns `B:D`, whole rows `2:5`)
 * (`Session.patchCells`): each field given sets the cells' own, null removes it, a field left out stays; values and
 * formulas stay as they are. A style by its fields (`StylePatch`) goes over the style each cell shows, its own or else
 * its row's or column's, as Excel formats a cell; a style by its number of the book's `styles` replaces it.
 */
export type CellPatch = { sheet: string; set: PatchFields } & ({ cell: string; range?: never } | { range: string; cell?: never })

/** The fields `CellPatch` sets; on whole rows and columns, `meta`, `note` and `link` can only be removed. */
export interface PatchFields {
  style?: number | StylePatch | null
  numberFormat?: NumberFormat | null
  /** @deprecated The name is `numberFormat`. */
  format?: NumberFormat | null
  meta?: JsonValue | null
  note?: CellNote | null
  link?: CellLink | null
}

/** Where a sheet's cells end (`Session.getUsedRange`): everything lies in its first `rows` rows and `columns` columns; 0 and 0 for an empty sheet. */
export interface SheetExtent {
  rows: number
  columns: number
}

/** A sheet without its cells (`Session.getLayout`). */
export type SheetLayout = Omit<SheetInput, 'cells'>

/** A book without the cells of its sheets (`Session.getLayout`): its styles, names and settings, and each sheet's layout. */
export type BookLayout = Omit<BookInput, 'sheets'> & { sheets: SheetLayout[] }

/** A session's book and its result taken at one moment (`Session.getSnapshot`), as `writeXlsx` takes them. */
export interface BookSnapshot {
  book: BookInput
  result: ComputeResult
}

/**
 * A session over a book: edits recompute only the cells that depend on them. The current clock of the
 * session (`clock`) is what a saved result should record: a full computation with it gives the same hash.
 */
export interface Session {
  readonly clock: Clock
  /** The book's revision: 0 when opened, one more after each call that changed the book (`EditResult.revision`). */
  readonly revision: number
  /**
   * Edits of cells, applied atomically with one recomputation (`CellEdit`): each sets the fields it gives and keeps the
   * cell's others, as Excel keeps a cell's look when a value is typed over it; null clears the content, as Delete. An
   * edit that keeps the cell's value and formula (a new `style`, `numberFormat`, `meta`, `note` or `link`) recomputes nothing.
   * A value or formula put into a spill range stops the array spilling (#SPILL!); clearing it lets the array spill
   * again. Edits that change some cells of an old-kind array formula's area (`CellInput.array`) but not all of them
   * throw `InputError` and change nothing, as Excel refuses to change part of an array; an edit of its top-left cell that
   * keeps an array formula over the same area changes all of it. An edit of a table's header cell renames its column, as
   * Excel does: the structured references to it follow, a number or a boolean becomes its text, a cleared cell names the
   * column `Column1` (the first such name free, in the word of the book's locale: `Столбец1` in ru-RU), and a name
   * another column has already gives the column further right a number (`Price2`). A formula in a header cell throws
   * `InputError`.
   */
  setCells(edits: CellEdit[], opts?: EditOptions): EditResult
  /** Cells replaced whole, as `setCells` applies edits otherwise: what an input leaves out the cell loses, null removes the cell. */
  replaceCells(edits: CellReplace[], opts?: EditOptions): EditResult
  /**
   * The cells of rectangles cleared, as Excel's ClearContents (`contents`: values, formulas and hyperlinks), ClearFormats
   * (`formats`: styles and number formats, the conditional formats over them) and Clear (`all`: both, notes, meta, the
   * rules of data validation over them); whole rows and columns lose their own style and format too (`docs/edits.md`).
   */
  clearCells(ranges: RangeRef[], what: ClearKind, opts?: EditOptions): EditResult
  /**
   * The look of cells and ranges (`CellPatch`), applied in order and atomically, without the cells' values and
   * formulas, which stay: an empty cell of a rectangle gets a cell with the look, a cell left with nothing is removed, a
   * cell of an old-kind array formula's area or of a spill takes it and the array stays. A style by its fields goes over
   * each cell's (`docs/edits.md`); equal styles share one entry of the book's `styles`. Recomputes nothing.
   */
  patchCells(patches: CellPatch[], opts?: EditOptions): EditResult
  /**
   * Text typed into a cell, as Excel takes it in the book's locale, date system and clock, by the format the cell shows
   * (`parseEntry`): the cell keeps its style, meta, note and link, and a format it shows from its row or column is not
   * written into it. A number that does not fit a column of the default width widens it in the same step (`width`), as
   * Excel does, unless `widen` is false; a value the cell's rule of data validation does not take is written, with the
   * rule (`validation`). A formula Excel does not take, or a change of part of an array, is refused, with a message in
   * the book's locale.
   */
  setCellText(sheet: string, cell: string, text: string, opts?: CellTextOptions): CellTextResult
  /** What `setCellText` would do, with nothing changed. */
  previewCellText(sheet: string, cell: string, text: string, opts?: CellTextOptions): CellTextPreview
  setNames(names: NameInput[], opts?: EditOptions): EditResult
  /** Non-live sessions only. */
  setClock(clock: Clock, opts?: EditOptions): EditResult
  /** A structure operation on the session's book, recomputed; refused as `editStructure` refuses it. A sheet hidden or shown recomputes nothing. */
  editStructure(op: StructureOp, opts?: EditOptions): SessionStructureResult
  /** Replaces the book's style table; every style that cells, rows and columns use must stay in it. Recomputes nothing; the shown text does not depend on styles. */
  setStyles(styles: CellStyle[], opts?: EditOptions): EditResult
  /**
   * Column widths, row heights, row and column styles and formats, hidden rows and columns, outline levels, merged
   * areas, the filter mode, frozen panes, conditional formats, rules of data validation and outline settings; applied
   * atomically. Merged areas stop arrays spilling over them: the formulas whose spill range an added or removed area
   * meets are recomputed. Hidden rows and the filter mode change what SUBTOTAL and AGGREGATE count: their formulas over
   * the rows concerned are recomputed. Nothing else is recomputed. With `display`, the cells whose shown text or colour
   * a width, a format or a merged area changed are listed too.
   */
  setLayout(edits: LayoutEdit[], opts?: EditOptions): EditResult
  /**
   * The session's book with every change so far: what a host saves as a version. Its objects are the
   * session's own and later calls change them, so do not modify it; serialize or copy what must stay.
   */
  getBook(): BookInput
  /**
   * The session's book without its cells: the styles, names and settings, and every field of each sheet but `cells`
   * (columns, rows, merged areas, tables, frozen panes, conditional formats, data validation and the rest). It costs
   * what the layout holds, not the cells; its objects are the session's own, as `getBook()`'s are.
   */
  getLayout(): BookLayout
  /**
   * Where a sheet's cells end: the last row and column that a cell of its input (a value, a formula, or a style, a
   * format, meta, a note or a link alone), a spilled value (the whole spill range), a merged area or a table reaches.
   * Settings of whole rows and columns, conditional formats and data validation do not count. It costs the sheet's
   * columns, spills and cells without a value, not a pass over all its cells.
   */
  getUsedRange(sheet: string): SheetExtent
  explain(sheet: string, cell: string, opts?: ExplainOptions): Explanation
  /**
   * The cells of a rectangle of a sheet (`B2:D9`, `B2`, whole columns `B:D`, whole rows `2:9`), or of several (a cell in
   * two of them is listed once), as `getResult()` gives them, by address: the same cells (spilled ones too, empty ones
   * left out) with the same fields, and each cell's own `style`, `note` and `link`, an old-kind array formula's `array`,
   * and the `appearance` the sheet's conditional formats give it by the current values (an empty cell with one is listed
   * with the value null). For a window over a sheet: it costs what the rectangles hold, not the book; the statistics of
   * a rule's cells (lowest, highest, percentiles) are kept until an edit changes them. Edits do not list cells whose
   * conditional appearance changed: they give the rules' cells as `repaint`, for a host to read again where its window meets them.
   */
  getCells(sheet: string, range: string | string[]): Record<string, CellView>
  /**
   * A window over a sheet for a grid (`B2:D9`, `B:D`, `2:9`): its cells with their styles and number formats resolved,
   * empty ones with a style of their row or column too, the widths and heights of its lines, the merged areas and tables
   * that meet it and the sheet's frozen panes (`Viewport`). It costs what the window holds, as `getCells`, not the book.
   */
  getViewport(sheet: string, range: string, opts?: ViewportOptions): Viewport
  /** Why a cell holds an error and the cell it began in (`CellIssueDetail`), without the book's result; null for a cell without one. */
  getCellIssue(sheet: string, cell: string): CellIssueDetail | null
  /** The data validation of a cell (`CellValidation`), its list computed now; null when no rule is over the cell. */
  getValidation(sheet: string, cell: string): CellValidation | null
  /**
   * Whether the cell holding `input` (null: cleared) would meet its rule, as Excel checks a cell's value
   * (`Range.Validation.Value`): the value the input computes to, in the book with it. Changes nothing; a cell without a
   * rule takes anything. Text typed into the cell is `parseEntry`'s input first.
   */
  validateEntry(sheet: string, cell: string, input: CellInput | null): EntryCheck
  getResult(): ComputeResult
  /** The book (as `getBook()`, the session's own objects) and its result (as `getResult()`) in one call. */
  getSnapshot(): BookSnapshot
  /** The values hash `getResult()` gives, without building the result. */
  getHash(): string
  /**
   * Takes back the last change of the book, as Excel's Undo: the whole of one call that changed it (`UndoCall`),
   * recomputed; null when there is nothing to undo. A change after an undo forgets what `redo` could do again.
   */
  undo(opts?: EditOptions): UndoResult | null
  /** Does again the last change `undo` took back; null when there is none. */
  redo(opts?: EditOptions): UndoResult | null
  /** The steps `undo` and `redo` can take now. */
  readonly history: SessionHistory
  close(): void
  /** @deprecated The name is `editStructure`. */
  restructure(op: StructureOp, opts?: EditOptions): SessionStructureResult
  /** @deprecated The name is `getBook`. */
  book(): BookInput
  /** @deprecated The name is `getLayout`. */
  layout(): BookLayout
  /** @deprecated The name is `getUsedRange`. */
  extent(sheet: string): SheetExtent
  /** @deprecated The name is `getCells`. */
  cells(sheet: string, range: string | string[]): Record<string, CellView>
  /** @deprecated The name is `getValidation`. */
  validation(sheet: string, cell: string): CellValidation | null
  /** @deprecated The name is `getResult`. */
  result(): ComputeResult
  /** @deprecated The name is `getHash`. */
  hash(): string
}

// ---------- asynchronous and threaded computation (`docs/async.md`) ----------

/** What a computation needs of an AbortSignal: the platform's `AbortSignal` is one. */
export interface AbortSignalLike {
  readonly aborted: boolean
  readonly reason?: unknown
  addEventListener(type: 'abort', listener: () => void, options?: { once?: boolean }): void
  removeEventListener(type: 'abort', listener: () => void): void
}

/**
 * Why an asynchronous call failed (`CalcError.code`). `soft_timeout`: its `timeoutMs` passed, seen by a check inside
 * the computation. `aborted`: its `signal` fired. `hard_timeout`: it ran past the pool's `hardTimeoutMs` and its
 * thread was terminated. `out_of_memory`: its thread reached `maxHeapMb`. `thread_failed`: its thread exited or threw
 * outside the computation. `restore_failed`: the session could not be restored on a thread and is closed. `limit`:
 * the book is over `CalcLimits`. `queue_full`: `maxQueue` tasks already wait. `too_many_sessions`: the pool holds
 * `maxSessions` sessions. `closed`: the calculator or the session is closed. `stale_revision`: an edit's `baseRevision`
 * is not the session's revision (`EditOptions`); it changed nothing.
 */
export type CalcErrorCode =
  | 'soft_timeout'
  | 'aborted'
  | 'hard_timeout'
  | 'out_of_memory'
  | 'thread_failed'
  | 'restore_failed'
  | 'limit'
  | 'queue_full'
  | 'too_many_sessions'
  | 'closed'
  | 'stale_revision'

/** Options of one asynchronous call. */
export interface CalcOptions {
  /** Stops the call: a waiting call at once, a running one at its next check. */
  signal?: AbortSignalLike
  /** A soft deadline: milliseconds of computing, counted from when the call starts computing (not while it waits). */
  timeoutMs?: number
}

export type AsyncComputeOptions = ComputeOptions & CalcOptions
/** `signal` and `timeoutMs` bound the opening (the first computation of the book). */
export type AsyncSessionOptions = SessionOptions & CalcOptions

/** `AsyncSession.editStructure`: as `SessionStructureResult`, without the book (`AsyncSession.getBook()` gives it). */
export type AsyncStructureResult = (EditResult & { ok: true; mapCell(ref: CellRef): CellRef | null }) | Refusal<StructureRefusal>

/** @deprecated The name is `AsyncStructureResult`. */
export type AsyncRestructureResult = AsyncStructureResult

/** The options of an asynchronous call that changes a session's book. */
export type AsyncEditOptions = CalcOptions & EditOptions

/**
 * A session computed in portions or in a thread: the methods of `Session`, as promises. Calls are taken in order,
 * each after the one before has finished. A call that fails (a deadline, an abort, a limit) changes nothing.
 */
export interface AsyncSession {
  /** The session's clock as of its last finished call. */
  readonly clock: Clock
  /** The session's revision as of its last finished call. */
  readonly revision: number
  setCells(edits: CellEdit[], opts?: AsyncEditOptions): Promise<EditResult>
  replaceCells(edits: CellReplace[], opts?: AsyncEditOptions): Promise<EditResult>
  clearCells(ranges: RangeRef[], what: ClearKind, opts?: AsyncEditOptions): Promise<EditResult>
  patchCells(patches: CellPatch[], opts?: AsyncEditOptions): Promise<EditResult>
  setCellText(sheet: string, cell: string, text: string, opts?: CalcOptions & CellTextOptions): Promise<CellTextResult>
  previewCellText(sheet: string, cell: string, text: string, opts?: CalcOptions & CellTextOptions): Promise<CellTextPreview>
  setNames(names: NameInput[], opts?: AsyncEditOptions): Promise<EditResult>
  setClock(clock: Clock, opts?: AsyncEditOptions): Promise<EditResult>
  editStructure(op: StructureOp, opts?: AsyncEditOptions): Promise<AsyncStructureResult>
  setStyles(styles: CellStyle[], opts?: AsyncEditOptions): Promise<EditResult>
  setLayout(edits: LayoutEdit[], opts?: AsyncEditOptions): Promise<EditResult>
  /** The session's book; from a thread, a copy. */
  getBook(): Promise<BookInput>
  /** As `Session.getLayout`; from a thread, a copy, without the cells. */
  getLayout(): Promise<BookLayout>
  getUsedRange(sheet: string): Promise<SheetExtent>
  explain(sheet: string, cell: string, opts?: ExplainOptions): Promise<Explanation>
  /** As `Session.getCells`; from a thread only the rectangles come back, in one call. */
  getCells(sheet: string, range: string | string[]): Promise<Record<string, CellView>>
  /** As `Session.getViewport`; from a thread the window comes back in one call. */
  getViewport(sheet: string, range: string, opts?: ViewportOptions): Promise<Viewport>
  getCellIssue(sheet: string, cell: string): Promise<CellIssueDetail | null>
  getValidation(sheet: string, cell: string): Promise<CellValidation | null>
  validateEntry(sheet: string, cell: string, input: CellInput | null, opts?: CalcOptions): Promise<EntryCheck>
  getResult(opts?: CalcOptions): Promise<ComputeResult>
  /** As `Session.getSnapshot`, in one call, so no edit comes between the book and its result; from a thread, a copy. */
  getSnapshot(opts?: CalcOptions): Promise<BookSnapshot>
  /** As `Session.getHash`: from a thread only the hash comes back. */
  getHash(opts?: CalcOptions): Promise<string>
  undo(opts?: AsyncEditOptions): Promise<UndoResult | null>
  redo(opts?: AsyncEditOptions): Promise<UndoResult | null>
  /** The session's history as of its last finished call. */
  readonly history: SessionHistory
  close(): Promise<void>
  /** @deprecated The name is `editStructure`. */
  restructure(op: StructureOp, opts?: AsyncEditOptions): Promise<AsyncStructureResult>
  /** @deprecated The name is `getBook`. */
  book(): Promise<BookInput>
  /** @deprecated The name is `getLayout`. */
  layout(): Promise<BookLayout>
  /** @deprecated The name is `getUsedRange`. */
  extent(sheet: string): Promise<SheetExtent>
  /** @deprecated The name is `getCells`. */
  cells(sheet: string, range: string | string[]): Promise<Record<string, CellView>>
  /** @deprecated The name is `getValidation`. */
  validation(sheet: string, cell: string): Promise<CellValidation | null>
  /** @deprecated The name is `getResult`. */
  result(opts?: CalcOptions): Promise<ComputeResult>
  /** @deprecated The name is `getHash`. */
  hash(opts?: CalcOptions): Promise<string>
}

/** A calculator: `createInlineCalc` (the same thread, in portions) or `createCalcPool` (threads). */
export interface Calc {
  compute(book: BookInput, opts?: AsyncComputeOptions): Promise<ComputeResult>
  createSession(book: BookInput, opts?: AsyncSessionOptions): Promise<AsyncSession>
  /** @deprecated The name is `createSession`. */
  openSession(book: BookInput, opts?: AsyncSessionOptions): Promise<AsyncSession>
  /** Refuses further calls (`closed`), the waiting and running ones too, and closes the sessions. */
  close(): Promise<void>
}

export interface InlineCalcOptions {
  /** Milliseconds of computing between two yields to the event loop; 10 by default. */
  sliceMs?: number
  /** The soft deadline of every call that gives none. */
  timeoutMs?: number
  /** The limits of every computation; a computation's own `limits` take their place field by field, the others stay. */
  limits?: CalcLimits
}

/** What the pool needs of a Web Worker (`CalcPoolOptions.spawn`). */
export interface WorkerLike {
  postMessage(message: unknown): void
  terminate(): void
  addEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { data?: unknown; message?: string }) => void): void
}

export interface CalcPoolOptions {
  /** Threads; 2 by default. */
  threads?: number
  /** Node: the heap of each thread (`resourceLimits.maxOldGenerationSizeMb`); not available in a browser. */
  maxHeapMb?: number
  /** A task that runs longer, from when it reaches its thread, terminates the thread (`hard_timeout`). */
  hardTimeoutMs?: number
  /** The soft deadline of every call that gives none. */
  timeoutMs?: number
  /** Tasks waiting for a thread, all threads together; one more is refused (`queue_full`). */
  maxQueue?: number
  /** Open sessions of the pool; one more is refused (`too_many_sessions`). */
  maxSessions?: number
  /** Sessions a thread holds; past it the session called least recently leaves the thread and comes back on its next call. */
  maxSessionsPerThread?: number
  /**
   * The limits of every computation; a computation's own `limits` take their place field by field, the others stay. With `maxHeapMb`, `cells` is `maxHeapMb · 2^20 / 64` where not
   * given, and `spill` the smaller of `cells` and `maxHeapMb · 2^20 / 1024` where not given: one allocation larger than
   * what a thread's heap has left (a large array, the tables of a large spill's cells) stops the whole process, not the
   * thread.
   */
  limits?: CalcLimits
  /** Calls kept in a session's journal before the pool takes a new snapshot of its book; 64 by default. */
  snapshotEvery?: number
  /** Milliseconds of computing between two yields inside a thread, where a cancel is read; 50 by default. */
  sliceMs?: number
  /** Starts a thread in a browser, for a bundler that serves the thread's script elsewhere. */
  spawn?: () => WorkerLike
}

/** @deprecated The package's named exports are these functions (`compute`, `createSession` and the rest): import them. */
export interface FormulaEngine {
  compute(book: BookInput, opts?: ComputeOptions): ComputeResult
  validateFormula(text: string, ctx?: ValidateContext): ValidateResult
  /** Excel's rules for a name (`NameInput.name`) in a book of `opts.locale` (en-US without one; a locale's words for TRUE and FALSE are refused too, ИСТИНА and ЛОЖЬ in ru-RU). */
  validateName(name: string, opts?: NameOptions): NameCheck
  /** The issues of a book's names, as `ComputeResult.diagnostics` of kind `name` give them for a book with these sheets, tables and locale. */
  validateNames(names: NameInput[], ctx: NamesContext): NameIssue[]
  createSession(book: BookInput, opts?: SessionOptions): Session
  editStructure(book: BookInput, op: StructureOp): StructureResult
  /** @deprecated The name is `createSession`. */
  openSession(book: BookInput, opts?: SessionOptions): Session
  /** @deprecated The name is `editStructure`. */
  restructure(book: BookInput, op: StructureOp): StructureResult
  /** A value as a cell shows it, without a book: by a format, in a locale and a column's width. */
  displayValue(value: CellValue, opts?: DisplayOptions): CellDisplay
  /** Text typed into a cell as Excel takes it, without a book: by the book's locale and the cell's format, style and meta. */
  parseEntry(text: string, opts?: EntryOptions): EntryResult
  /**
   * The widths AutoFit gives columns `C` or `B:D` of a computed book's sheet, from the values of `result`, by column
   * letter: null for the default width; a column with nothing to show is left out, as it keeps its width.
   */
  fitColumns(book: BookInput, result: ComputeResult, sheet: string, columns: string): Record<string, number | null>
  /**
   * The width a column takes as a value typed into one of its cells (the value it computes, after `setCells`) widens
   * it, as Excel widens a column whose width is the default or one it fitted for a number AutoFit would give more room;
   * undefined when the column keeps its width. A host sets it with `bestFit: true`.
   */
  entryWidth(value: CellValue, opts?: EntryWidthOptions): number | undefined
}

export const MAX_FORMULA_LENGTH = 8_192
export const MAX_FORMULA_DEPTH = 64
