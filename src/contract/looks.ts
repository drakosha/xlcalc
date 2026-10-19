/** Looks and rules: the categories of number formats, styles with their fonts, borders and alignments, a style's fields to change, conditional formats and what they give a cell, rules of data validation. */
import type { NumberFormat } from './book.js'

/**
 * A category of number formats as the Number Format box of Excel's Home tab shows it for the active cell
 * (`numberFormatCategory`): the category of Excel's Format Cells window that makes the format, or `custom` for a format
 * none of them makes.
 */
export type FormatCategory = 'general' | 'number' | 'currency' | 'accounting' | 'date' | 'time' | 'percentage' | 'fraction' | 'scientific' | 'text' | 'special' | 'custom'

/** An item of the Number Format box's list on the Home tab, in its order (`numberFormatChoices`). */
export type FormatChoiceKey = 'general' | 'number' | 'currency' | 'accounting' | 'shortDate' | 'longDate' | 'time' | 'percentage' | 'fraction' | 'scientific' | 'text'

/** What the Number Format box shows for a format: its category and the category's name in the locale (`Number`, `Числовой`). */
export interface NumberFormatCategory {
  category: FormatCategory
  name: string
}

/**
 * An item of the Number Format list in a locale: its key, its name in the list (`Short Date`), the format it applies (a
 * built-in format by number, or a code as Excel writes it to a file) and that format's code in the locale, in en-US
 * syntax as a cell's `numberFormat` takes it.
 */
export interface FormatChoice {
  key: FormatChoiceKey
  name: string
  format: NumberFormat
  code: string
}

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

/** A cell's alignment (`CellStyle.alignment`), as xlsx `<alignment>`: horizontal and vertical, text wrapped, indent. */
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

/** What every kind of conditional format shares (`ConditionalFormat`): the cells it applies to and Excel's Stop If True. */
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

/** Fields of a style to change (`CellPatch.set.style`): a value sets a field, null removes it, a field left out stays. */
export interface StylePatch {
  font?: { [K in keyof FontStyle]?: FontStyle[K] | null } | null
  fill?: { color: Color } | null
  border?: { [K in keyof NonNullable<CellStyle['border']>]?: BorderLine | null } | null
  alignment?: { [K in keyof AlignmentStyle]?: AlignmentStyle[K] | null } | null
}
