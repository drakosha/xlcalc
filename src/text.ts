/**
 * Formula text without the engine (the `./text` subpath, the same functions in the root): the references of text as it
 * is typed, for colouring them, what an arrow does in it (Point mode, `point.ts`), a sheet's name and a table column's name as a reference writes them, the clipboard's
 * text as the rectangle Excel pastes, and the category of a number format as the ribbon's Number Format box shows it,
 * with the items of its list (`numfmt/category.ts`, by the engine's reader of format codes). The tokenizer, the parser
 * and the translation of a locale's formula text are the engine's own (`scan.ts`, `locale/translate.ts`); they compute nothing,
 * and no locale's data comes with them but en-US's messages, which the tokenizer's errors carry: another locale is given
 * as an object (`ruRU` of `xlcalc/locales/ru-RU`). Each function refuses options it does not take (`strict-args.ts`).
 */
import { checkCellValue } from './input.js'
import { InputError, what } from './input-error.js'
import { array, checkBoolean, checkNumber, checkString, object, onlyOptions } from './input/basic.js'
import type { CellValue, CompareValuesOptions, FindTextOptions, NumberFormat } from './contract.js'
import { checkFormat } from './input/styles.js'
import { clipboardRows } from './entry/clipboard.js'
import { folded, matcher, replacer } from './find-text.js'
import { sameType, sortOrder } from './sort-order.js'
import { textOrder } from './collate.js'
import { optionalLocale } from './locale/load.js'
import { charsOfPoints, columnPoints, MAX_POINTS, pointsOfChars, widthOf } from './column-width.js'
import * as call from './call-at.js'
import * as category from './numfmt/category.js'
import * as decimals from './numfmt/decimals.js'
import * as point from './point.js'
import * as scan from './scan.js'
import { strict } from './strict-args.js'
import { escapeColumn as escaped } from './syntax/table.js'

export { quoteSheetIfNeeded } from './a1.js'
export { DEFAULT_ROW_HEIGHT, MAX_OUTLINE, MAX_ROW_HEIGHT } from './column-width.js'
export { FUNCTION_CATEGORIES } from './function-categories.js'

export const scanReferences = /* @__PURE__ */ strict(scan.scanReferences)
export const refAt = /* @__PURE__ */ strict(point.refAt)
export const callAt = /* @__PURE__ */ strict(call.callAt)
export const nameAt = /* @__PURE__ */ strict(call.nameAt)
export const numberFormatCategory = /* @__PURE__ */ strict(category.numberFormatCategory)
export const numberFormatChoices = /* @__PURE__ */ strict(category.numberFormatChoices)

/**
 * The number format Excel's Increase Decimal (`step` 1) or Decrease Decimal (-1) gives a cell of `format` (General
 * without one) holding `value`, in `opts.locale` (a built-in format's code in it) and a column of `opts.width` (Excel's
 * standard without one), as Excel for Mac does: a General cell of a number takes its decimals one more or one less
 * (`0.00` for 1.5), all of them when General shows the number whole, else those it shows in the column; a code takes a
 * decimal more or less in each section of digits, a copy of its last one (`#.##` to `#.###`); a date, a time, a
 * fraction, text, accounting's zero (`"-"??`) stay. It returns `format` itself when the step changes nothing: no number
 * in General, no decimal to take away. Excel applies the active cell's result to the whole selection.
 */
export const stepDecimals = /* @__PURE__ */ strict(function stepDecimals(format: NumberFormat | undefined, value: CellValue, step: 1 | -1, opts?: { locale?: string; width?: number }): NumberFormat | undefined {
  if (format !== undefined) checkFormat(format, 'format')
  if (step !== 1 && step !== -1) throw new InputError('step', `expected 1 or -1, got ${what(step)}`)
  const o = opts === undefined ? {} : object(opts, 'opts', "options such as { locale: 'ru-RU', width: 12 }")
  onlyOptions(o, 'opts', ['locale', 'width'])
  optionalLocale(o.locale, 'opts.locale')
  if (o.width !== undefined) checkNumber(o.width, 'opts.width', 0, 255)
  return decimals.stepDecimals(format, checkCellValue(value, 'value'), step, o.locale as string | undefined, o.width as number | undefined)
})

/**
 * A column's name in a structured reference as Excel writes it: `'` before each `[`, `]`, `#` and `'` in it, and before
 * a leading `@` (`x#y` becomes `x'#y`, `@at` becomes `'@at`), everything else as is (`Unit Price`).
 */
export const escapeColumn = /* @__PURE__ */ strict(function escapeColumn(name: string): string {
  return escaped(checkString(name, 'name'))
})

/**
 * Text from the clipboard as the rectangle Excel pastes (`Session.pasteText`): rows split at line breaks, made of fields
 * split at tabs, each row as wide as the longest; a field that begins with a quote runs to its closing quote, with line
 * breaks and tabs inside it, and a doubled quote stands for one; an empty field clears its cell, and a line break after
 * the last row is a row of empty fields.
 */
export const parseClipboard = /* @__PURE__ */ strict(function parseClipboard(text: string): string[][] {
  return clipboardRows(checkString(text, 'text'))
})

/** Options of boolean flags checked: an object of `keys` and `more` alone, each of `keys` a boolean (`hint`: an example for the message). */
function flags<K extends string>(opts: object | undefined, keys: readonly K[], hint: string, more: readonly string[] = []): Partial<Record<K, boolean>> {
  const o = opts === undefined ? {} : object(opts, 'opts', hint)
  onlyOptions(o, 'opts', [...keys, ...more])
  for (const k of keys) if (o[k] !== undefined) checkBoolean(o[k], `opts.${k}`)
  return o as Partial<Record<K, boolean>>
}

/** Excel's plain-text copy: tabs, CRLF rows and quoted control characters; Windows ends with CRLF, Mac does not. */
export const formatClipboard = /* @__PURE__ */ strict(function formatClipboard(rows: string[][], opts?: { trailingNewline?: boolean }): string {
  const o = flags(opts, ['trailingNewline'], 'clipboard options such as { trailingNewline: false }')
  const grid = array(rows, 'rows', 'an array of rows of text')
  let width = 0
  const lines = grid.map((row, i) => {
    const at = `rows[${i}]`
    const fields = array(row, at, 'an array of cell texts')
    if (!fields.length || (i > 0 && fields.length !== width)) throw new InputError(at, 'expected a non-empty row as wide as the first row')
    width = fields.length
    return fields.map((field, j) => {
      const text = checkString(field, `${at}[${j}]`)
      return /[\t\r\n]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text
    }).join('\t')
  })
  return lines.join('\r\n') + (lines.length && o.trailingNewline !== false ? '\r\n' : '')
})

/** The widest stored width Excel for Mac writes, 255 digits and the padding (`MAX_POINTS` its points). */
const MAX_FILE_WIDTH = 255.83203125

/** A stored xlsx width in the library's Mac point model; null or undefined is the default 53 points. */
export const columnWidthToPoints = /* @__PURE__ */ strict(function columnWidthToPoints(width?: number | null): number {
  if (width !== undefined && width !== null) checkNumber(width, 'width', 0, MAX_FILE_WIDTH)
  return columnPoints(width ?? undefined)
})

/** Whole points back to a stored width in 256ths; fractional points round down, as the model's geometry does. */
export const columnWidthFromPoints = /* @__PURE__ */ strict(function columnWidthFromPoints(points: number): number {
  checkNumber(points, 'points', 0, MAX_POINTS)
  return widthOf(Math.trunc(points))
})

/**
 * A stored xlsx width as Excel for Mac's Column Width box displays it, in characters of the workbook's default font, to
 * two decimal places (the default column, null or undefined, is 8); in the library's Mac point model, like
 * `columnWidthToPoints`.
 */
export const columnWidthToChars = /* @__PURE__ */ strict(function columnWidthToChars(width?: number | null): number {
  if (width !== undefined && width !== null) checkNumber(width, 'width', 0, MAX_FILE_WIDTH)
  return charsOfPoints(columnPoints(width ?? undefined))
})

/** Converts a width in characters, as typed into Excel for Mac's Column Width box (0 to 255), to the stored xlsx width Excel makes of it. */
export const columnWidthFromChars = /* @__PURE__ */ strict(function columnWidthFromChars(chars: number): number {
  checkNumber(chars, 'chars', 0, 255)
  return widthOf(pointsOfChars(chars))
})

const FIND = ['caseSensitive', 'wildcards', 'wholeCell'] as const

/** Excel's Find over a host's own texts, matching as `findCells` matches a cell's text: a test for a text, or null when nothing can match (an empty query, `~` alone). */
export const findMatcher = /* @__PURE__ */ strict(function findMatcher(query: string, opts?: FindTextOptions): ((text: string) => boolean) | null {
  const o = flags(opts, FIND, 'Find options such as { wildcards: true }')
  const test = matcher(checkString(query, 'query'), o)
  return test && ((text) => test(o.caseSensitive ? text : folded(text)))
})

/** Excel's Replace over a host's own texts, as `replaceInCells` replaces in a cell's text: the text with each match replaced by `by`, or null where nothing matches. */
export const findReplacer = /* @__PURE__ */ strict(function findReplacer(query: string, by: string, opts?: FindTextOptions): ((text: string) => string | null) | null {
  return replacer(checkString(query, 'query'), checkString(by, 'by'), flags(opts, FIND, 'Find options such as { wildcards: true }'))
})

/** The order of values of one type by a locale's order of text (`RegionalData.collation`), each made once. */
const ORDERS = new Map<string, (a: CellValue, b: CellValue) => number>()

/** The options of the comparisons, checked: the reversal, and the order of one type (en-US's, `sortOrder`'s own, as undefined). */
function compareOptions(opts: CompareValuesOptions | undefined): [boolean, ((a: CellValue, b: CellValue) => number) | undefined] {
  const desc = flags(opts, ['descending'], 'options such as { descending: true }', ['locale']).descending === true
  const tag = optionalLocale(opts?.locale, 'opts.locale')?.regional?.collation ?? 'en-US'
  if (tag === 'en-US') return [desc, undefined]
  let same = ORDERS.get(tag)
  if (!same) ORDERS.set(tag, (same = sameType(textOrder(tag))))
  return [desc, same]
}

/**
 * Excel's sort order for two cells' values, as the Sort command of a host's display filter sorts them: numbers, then
 * text in any case (in `locale`'s text order, en-US's without it), then FALSE and TRUE, then errors, all equal to each
 * other, and blanks (null) last, even with `descending`; negative when `a` comes first, 0 for equal values, which a
 * stable sort keeps in their original order, as Excel does. The SORT function instead orders errors by their
 * ERROR.TYPE number.
 */
export const compareValues = /* @__PURE__ */ strict(function compareValues(a: CellValue, b: CellValue, opts?: CompareValuesOptions): number {
  return sortOrder(checkCellValue(a, 'a'), checkCellValue(b, 'b'), ...compareOptions(opts))
})

/**
 * `compareValues` with its options checked once, for `Array.prototype.sort` over many values: the order of two cells'
 * values, which it accepts as `CellValue`s without checking them (`values.sort(valueComparator({ descending: true }))`).
 */
export const valueComparator = /* @__PURE__ */ strict(function valueComparator(opts?: CompareValuesOptions): (a: CellValue, b: CellValue) => number {
  const [desc, same] = compareOptions(opts)
  return (a, b) => sortOrder(a, b, desc, same)
})
