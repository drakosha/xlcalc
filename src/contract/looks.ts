/** Looks and rules: the categories of number formats, styles with their fonts, borders and alignments, a style's fields to change, conditional formats and what they give a cell, rules of data validation. */
import type { NumberFormat } from './book.js'

/**
 * A number format category, as the Number Format box on Excel's Home tab displays it for the active cell
 * (`numberFormatCategory`): the category of Excel's Format Cells dialog that produces the format, or `custom` for a
 * format that none of them produces.
 */
export type FormatCategory = 'general' | 'number' | 'currency' | 'accounting' | 'date' | 'time' | 'percentage' | 'fraction' | 'scientific' | 'text' | 'special' | 'custom'

/** An item in the Number Format box's list on the Home tab, in list order (`numberFormatChoices`). */
export type FormatChoiceKey = 'general' | 'number' | 'currency' | 'accounting' | 'shortDate' | 'longDate' | 'time' | 'percentage' | 'fraction' | 'scientific' | 'text'

/** What the Number Format box displays for a format: its category and the category's name in the locale (`Number`, `Числовой`). */
export interface NumberFormatCategory {
  category: FormatCategory
  name: string
}

/**
 * An item of the Number Format list in a locale: its key, its name in the list (`Short Date`), the format it applies (a
 * built-in format by number, or a format code as Excel writes it to a file), and that format's code in the locale, in
 * en-US syntax as a cell's `numberFormat` accepts it.
 */
export interface FormatChoice {
  key: FormatChoiceKey
  name: string
  format: NumberFormat
  code: string
}

/** A color, `#RRGGBB`. */
export type Color = string

/** A font; whatever a style omits comes from the workbook's default font. */
export interface FontStyle {
  name?: string
  /** In points, 1 to 409. */
  size?: number
  bold?: boolean
  italic?: boolean
  underline?: 'single' | 'double'
  strike?: boolean
  color?: Color
}

/** One side of a cell's border; the line styles are those of xlsx. */
export interface BorderLine {
  style: 'thin' | 'medium' | 'thick' | 'dotted' | 'dashed' | 'double' | 'hair' | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot'
  color?: Color
}

/** A cell's alignment (`CellStyle.alignment`), like xlsx `<alignment>`: horizontal and vertical alignment, text wrapping, indent. */
export interface AlignmentStyle {
  horizontal?: 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed'
  vertical?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed'
  wrap?: boolean
  /** Indent level, 0 to 250. */
  indent?: number
}

/**
 * A cell style in the workbook's style table (`BookInput.styles`), like an xlsx cell format without its number format
 * (which is the cell's `numberFormat`). Styles play no part in calculation or in the values hash.
 */
export interface CellStyle {
  font?: FontStyle
  /** A solid fill. */
  fill?: { color: Color }
  border?: { top?: BorderLine; right?: BorderLine; bottom?: BorderLine; left?: BorderLine }
  alignment?: AlignmentStyle
}

/**
 * A conditional format: the cells it applies to, a condition, and the formatting of a cell that meets it. The rules
 * whose conditions are true for a cell apply in the sheet's order: a style field set by an earlier rule wins, and
 * `stopIfTrue` skips the rules after one whose condition is true. Its formulas are stored text, like a cell's: a
 * leading `=` is removed.
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

/** What all kinds of conditional format have in common (`ConditionalFormat`): the cells it applies to and Excel's Stop If True. */
export interface ConditionalBase {
  /**
   * The cells it applies to: rectangles on the sheet (`B2:B20`, `D2`, entire columns `F:F`, entire rows `3:5`), like
   * Excel's "Applies to" box and xlsx `sqref`. Its formulas are written for the top-left cell of the rectangle enclosing
   * them: a reference without `$` shifts from there to each cell, as in a copied formula.
   */
  ranges: string[]
  /** Excel's Stop If True: the rules after this one do not apply to a cell for which this rule's condition is true. */
  stopIfTrue?: boolean
}

/** The formatting a rule applies to the cells that meet its condition, like an xlsx differential format (`dxf`): each field it sets replaces the cell's own. */
export interface ConditionalStyle {
  /** `bold: false` removes bold from a cell that has it, and so on. */
  font?: Pick<FontStyle, 'bold' | 'italic' | 'underline' | 'strike' | 'color'>
  fill?: { color: Color }
  border?: CellStyle['border']
  /** The number format used to display the cell's value. */
  numberFormat?: NumberFormat
}

/**
 * The cell's value compared with the value of one formula, or two for `between` and `notBetween` (both bounds
 * inclusive, in either order), the way `=`, `<` and the other operators compare; a blank cell that a formula refers to
 * counts as 0 there.
 */
export interface CellIsFormat extends ConditionalBase {
  type: 'cellIs'
  operator: 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual' | 'between' | 'notBetween'
  /** Formulas as a cell stores them, without `=`: `100`, `"done"`, `$B$1`. */
  formulas: string[]
  style?: ConditionalStyle
}

/** A formula whose result is true, as IF evaluates its condition: TRUE, a number other than 0, or the text TRUE. */
export interface ExpressionFormat extends ConditionalBase {
  type: 'expression'
  formula: string
  style?: ConditionalStyle
}

/**
 * The cell's text contains `text`, does not contain it, begins with it or ends with it, in any case. Excel calculates
 * the rule's formula, which it writes alongside the text (`NOT(ISERROR(SEARCH("ab",A1)))`); without `formula`, the
 * rule uses that formula.
 */
export interface TextFormat extends ConditionalBase {
  type: 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith'
  text: string
  formula?: string
  style?: ConditionalStyle
}

/**
 * The cell is blank (or contains only spaces), is not blank, contains an error, or does not. Blanks are determined by
 * the value, as Excel displays it (an error is not blank); errors by `formula`, or by Excel's (`ISERROR(A1)`) without
 * it. Excel writes a formula alongside both kinds of rule.
 */
export interface BlanksFormat extends ConditionalBase {
  type: 'containsBlanks' | 'notContainsBlanks' | 'containsErrors' | 'notContainsErrors'
  formula?: string
  style?: ConditionalStyle
}

/** The cell contains a date in the period, according to the workbook's clock (`TODAY()`); calculated by `formula`, or by Excel's formula without it. */
export interface TimePeriodFormat extends ConditionalBase {
  type: 'timePeriod'
  period: 'today' | 'yesterday' | 'tomorrow' | 'last7Days' | 'thisWeek' | 'lastWeek' | 'nextWeek' | 'thisMonth' | 'lastMonth' | 'nextMonth'
  formula?: string
  style?: ConditionalStyle
}

/** The `rank` highest numbers among the cells (the lowest with `bottom`), or `rank` percent of them with `percent`. */
export interface TopFormat extends ConditionalBase {
  type: 'top10'
  /** 1 to 1,000; 1 to 100 with `percent`. */
  rank: number
  percent?: boolean
  bottom?: boolean
  style?: ConditionalStyle
}

/** Numbers above the average of the cells' numbers (below it with `below`, or also equal to it with `equal`), or `stdDev` standard deviations beyond it. */
export interface AverageFormat extends ConditionalBase {
  type: 'aboveAverage'
  below?: boolean
  equal?: boolean
  /** 1 to 3. */
  stdDev?: number
  style?: ConditionalStyle
}

/** Values that occur more than once among the cells (text compared case-insensitively), or only once. */
export interface DuplicatesFormat extends ConditionalBase {
  type: 'duplicateValues' | 'uniqueValues'
  style?: ConditionalStyle
}

/**
 * A threshold of a color scale, a data bar or an icon set, over the numbers in the cells the rule applies to. `min`
 * and `max`: the lowest and the highest number; `num`: the number `value` returns; `percent`: `value` percent of the
 * way from the lowest to the highest; `percentile`: the `value`-th percentile of the numbers; `formula`: the result of
 * `value`. A data bar's `autoMin` and `autoMax` are Excel 2010's automatic bounds: 0, or the lowest number when it is
 * below 0 (the highest when it is above 0).
 */
export interface ConditionalValue {
  type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula' | 'autoMin' | 'autoMax'
  /** A formula as a cell stores it (`10`, `$B$1*2`), for `num`, `percent`, `percentile` and `formula`. */
  value?: string
}

/** Each number gets a color between those of the points on either side of it; other cells get none. */
export interface ColorScaleFormat extends ConditionalBase {
  type: 'colorScale'
  /** Two or three points, from the lowest. */
  points: { value: ConditionalValue; color: Color }[]
}

/**
 * Each number gets a bar across its cell, its length proportional to where the number lies between `min` and `max`.
 * The fields from `gradient` onward are Excel 2010's (xlsx `x14:dataBar`).
 */
export interface DataBarFormat extends ConditionalBase {
  type: 'dataBar'
  min: ConditionalValue
  max: ConditionalValue
  color: Color
  /** False: the cell displays the bar without its value. */
  showValue?: boolean
  /** The shortest and the longest bar as a percentage of the cell's width: 0 and 100 by default, 10 and 90 for an Excel 2007 bar. */
  minLength?: number
  maxLength?: number
  /** A gradient fill (true by default) or a solid one. */
  gradient?: boolean
  /** The bar's border color; no border without it. */
  border?: Color
  /** The direction of the bars; by default the sheet's direction (left to right). */
  direction?: 'leftToRight' | 'rightToLeft'
  /** The fill color of bars for negative numbers; the bar's color without it. */
  negativeColor?: Color
  /** The border color of those bars; `border` without it. */
  negativeBorder?: Color
  /** Where the axis between negative and positive bars is: placed according to the numbers (the default), at the middle of the cell, or none (negative bars run in the same direction as positive ones). */
  axis?: 'middle' | 'none'
  /** The axis color; black without it. */
  axisColor?: Color
  /** An Excel 2007 bar (from a file without the Excel 2010 part of the rule): a gradient without a border, negative numbers as shorter bars, no axis; the fields from `gradient` onward do not apply. */
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

/** An icon from a set: the set and the icon's index in it, from 0, in Excel's order (xlsx `iconId`: for `3Arrows`, 0 is the red down arrow and 2 the green up arrow). */
export interface Icon {
  set: IconSetName
  index: number
}

/** Each number gets an icon from the set according to the thresholds it reaches. */
export interface IconSetFormat extends ConditionalBase {
  type: 'iconSet'
  icons: IconSetName
  /**
   * One threshold per icon of the set, from the lowest; a number gets the icon of the last threshold it reaches (`gte`:
   * at or above it, true by default; false: above it), and always reaches the first one. `icon` substitutes a different
   * icon, `null` no icon (Excel 2010's custom icons).
   */
  points: { value: ConditionalValue; gte?: boolean; icon?: Icon | null }[]
  /** The icons in reverse order: the last one for the lowest numbers. */
  reverse?: boolean
  /** False: the cell displays the icon without its value. */
  showValue?: boolean
}

/**
 * The effect of a sheet's conditional formats on a cell (`CellView.appearance`), based on the cells' current values. A
 * field is set only where the cell's display differs from what its own style and value produce.
 */
export interface ConditionalAppearance {
  /** The styles of the rules whose conditions are true, merged in order (the first rule's field wins); a color scale's color is in `fill`. */
  style?: ConditionalStyle
  bar?: BarLook
  icon?: Icon
  /** A data bar or an icon set hides the value: the cell displays only its bar or icon. */
  hideValue?: true
}

/** A data bar in a cell: where it is drawn, as fractions of the cell's width from the side where its direction starts. */
export interface BarLook {
  /** From 0 to 1; `start` < `end`. */
  start: number
  end: number
  color: Color
  gradient: boolean
  border?: Color
  /** The axis between negative and positive bars, as a fraction of the width, for a rule with negative numbers. */
  axis?: number
  axisColor?: Color
  /** The bars run from right to left: the fractions are measured from the cell's right edge. */
  rightToLeft?: true
}

/**
 * What a data validation rule allows in a cell: any value (`any`, a rule that only shows an input message), a whole
 * number, a decimal number, a date or a time compared with the rule's formulas, text whose length is compared the same
 * way, an item of a list, or a value for which a formula (`custom`) returns TRUE.
 */
export type ValidationType = 'any' | 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | 'list' | 'custom'

/** How a value is compared with `formula1` (and `formula2`): between and not between use both, inclusive. */
export type ValidationOperator = 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual'

/**
 * A data validation rule, like xlsx `<dataValidation>`. Its formulas are stored formula text without `=` (en-US, like a
 * cell's; a leading `=` is removed), relative to the rule's first cell, the top-left cell of the first range in
 * `range`: a rule over `B2:B9` with `formula1: 'A2>0'` checks `A5>0` for B5. A literal list is a text constant of
 * comma-separated items (`'"Yes,No,Maybe"'`), as a file stores it in every locale; a list from cells is a reference
 * (`'$D$2:$D$9'`, `"'Lists'!$A$1:$A$20"`), a defined name, or a formula that returns a reference.
 */
export interface DataValidation {
  /** The cells: space-separated ranges (`B2:B20 D2`, `C:C`), like xlsx `sqref`; no other rule on the sheet overlaps them. */
  range: string
  type: ValidationType
  /** For whole, decimal, date, time and textLength: how the value is compared; `between` by default. */
  operator?: ValidationOperator
  /** The value, the lower bound, the list or the formula of a custom rule; every type except `any` has one. */
  formula1?: string
  /** The upper bound for `between` and `notBetween`. */
  formula2?: string
  /**
   * A blank cell passes, and so does any value while a bound or a list refers to a blank cell (Excel's Ignore blank,
   * xlsx `allowBlank`); true by default, as when Excel creates a rule.
   */
  allowBlank?: boolean
  /** A list displays its dropdown in the cell; true by default (false is xlsx `showDropDown="1"`). */
  dropdown?: boolean
  /** The input message Excel displays while the cell is selected: its title (Excel allows 32 characters) and text (255). */
  promptTitle?: string
  prompt?: string
  /** Whether the input message is displayed; true by default, as when Excel creates a rule. */
  showPrompt?: boolean
  /** The alert for a value the rule does not allow: stop (the default), warning or information, and its title (32) and text (225). */
  errorStyle?: 'stop' | 'warning' | 'information'
  errorTitle?: string
  error?: string
  /** Whether the alert is displayed; true by default, as when Excel creates a rule. Without it, Excel accepts any typed value. */
  showError?: boolean
}

/** Fields of a style to change (`CellPatch.set.style`): a value sets a field, null removes it, and an omitted field is left unchanged. */
export interface StylePatch {
  font?: { [K in keyof FontStyle]?: FontStyle[K] | null } | null
  fill?: { color: Color } | null
  border?: { [K in keyof NonNullable<CellStyle['border']>]?: BorderLine | null } | null
  alignment?: { [K in keyof AlignmentStyle]?: AlignmentStyle[K] | null } | null
}
