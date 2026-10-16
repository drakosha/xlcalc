/**
 * Formula text without the engine (the `./text` subpath, the same functions in the root): the references of text as it
 * is typed, for colouring them, what an arrow does in it (Point mode, `point.ts`), a sheet's name and a table column's name as a reference writes them, the clipboard's
 * text as the rectangle Excel pastes, and the category of a number format as the ribbon's Number Format box shows it,
 * with the items of its list (`numfmt/category.ts`, by the engine's reader of format codes). The tokenizer, the parser
 * and the translation of a locale's formula text are the engine's own (`scan.ts`, `locale/translate.ts`); they compute nothing,
 * and no locale's data comes with them but en-US's messages, which the tokenizer's errors carry: another locale is given
 * as an object (`ruRU` of `xlcalc/locales/ru-RU`). Each function refuses options it does not take (`strict-args.ts`).
 */
import { array, checkBoolean, checkCellValue, checkNumber, checkString, InputError, object, onlyOptions } from './input.js'
import type { CellValue, FindTextOptions } from './contract.js'
import { clipboardRows } from './entry/clipboard.js'
import { folded, matcher, replacer } from './find-text.js'
import { sortOrder } from './sort-order.js'
import { charsOfPoints, columnPoints, pointsOfChars, widthOf } from './column-width.js'
import * as call from './call-at.js'
import * as category from './numfmt/category.js'
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
 * A column's name in a structured reference as Excel writes it: `'` before each `[`, `]`, `#` and `'` in it and before
 * an `@` it starts with (`x#y` is `x'#y`, `@at` is `'@at`), the rest as it is (`Unit Price`).
 */
export const escapeColumn = /* @__PURE__ */ strict(function escapeColumn(name: string): string {
  return escaped(checkString(name, 'name'))
})

/**
 * Text from the clipboard as the rectangle Excel pastes (`Session.pasteText`): its rows, at line breaks, of its fields, at
 * tabs, each row as wide as the longest; a field begun by a quote runs to its closing quote, line breaks and tabs inside
 * it, a doubled quote one; an empty field clears its cell, and a line break after the last row is a row of empty fields.
 */
export const parseClipboard = /* @__PURE__ */ strict(function parseClipboard(text: string): string[][] {
  return clipboardRows(checkString(text, 'text'))
})

/** Options of boolean flags checked: an object of `keys` alone, each a boolean (`hint`: an example for the message). */
function flags<K extends string>(opts: Partial<Record<K, boolean>> | undefined, keys: readonly K[], hint: string): Partial<Record<K, boolean>> {
  const o = opts === undefined ? {} : object(opts, 'opts', hint)
  onlyOptions(o, 'opts', keys)
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

/** A stored xlsx width in the library's Mac point model; null or undefined is the default 53 points. */
export const columnWidthToPoints = /* @__PURE__ */ strict(function columnWidthToPoints(width?: number | null): number {
  if (width !== undefined && width !== null) checkNumber(width, 'width', 0, 255.83203125)
  return columnPoints(width ?? undefined)
})

/** Whole points back to a stored width in 256ths; fractional points round down, as the model's geometry does. */
export const columnWidthFromPoints = /* @__PURE__ */ strict(function columnWidthFromPoints(points: number): number {
  checkNumber(points, 'points', 0, 1535)
  return widthOf(Math.trunc(points))
})

/**
 * A stored xlsx width as Excel for Mac's Column Width box shows it, in characters of the book's default font to two
 * decimals (the default column, null or undefined, is 8); in the library's Mac point model, as `columnWidthToPoints`.
 */
export const columnWidthToChars = /* @__PURE__ */ strict(function columnWidthToChars(width?: number | null): number {
  if (width !== undefined && width !== null) checkNumber(width, 'width', 0, 255.83203125)
  return charsOfPoints(columnPoints(width ?? undefined))
})

/** A width in characters, as typed into Excel for Mac's Column Width box (0 to 255), to the stored xlsx width Excel makes of it. */
export const columnWidthFromChars = /* @__PURE__ */ strict(function columnWidthFromChars(chars: number): number {
  checkNumber(chars, 'chars', 0, 255)
  return widthOf(pointsOfChars(chars))
})

const FIND = ['caseSensitive', 'wildcards', 'wholeCell'] as const

/** Excel's Find over a host's texts, as `findCells` matches a cell's: the test of a text, null when nothing can match (an empty query, `~` alone). */
export const findMatcher = /* @__PURE__ */ strict(function findMatcher(query: string, opts?: FindTextOptions): ((text: string) => boolean) | null {
  const o = flags(opts, FIND, 'Find options such as { wildcards: true }')
  const test = matcher(checkString(query, 'query'), o)
  return test && ((text) => test(o.caseSensitive ? text : folded(text)))
})

/** Excel's Replace over a host's texts, as `replaceInCells` replaces in a cell's: a text with each match put as `by`, null where nothing matches. */
export const findReplacer = /* @__PURE__ */ strict(function findReplacer(query: string, by: string, opts?: FindTextOptions): ((text: string) => string | null) | null {
  return replacer(checkString(query, 'query'), checkString(by, 'by'), flags(opts, FIND, 'Find options such as { wildcards: true }'))
})

/** The `descending` option of the comparisons, checked. */
const descending = (opts: { descending?: boolean } | undefined) => flags(opts, ['descending'], 'options such as { descending: true }').descending === true

/**
 * Excel's sort order of two cells' values, for a host's display filter: numbers, text in any case (en-US's collation),
 * FALSE then TRUE, errors (equal), and blanks (null) last, also with `descending`; negative when `a` comes first.
 */
export const compareValues = /* @__PURE__ */ strict(function compareValues(a: CellValue, b: CellValue, opts?: { descending?: boolean }): number {
  return sortOrder(checkCellValue(a, 'a'), checkCellValue(b, 'b'), descending(opts))
})

/**
 * `compareValues` with its options checked once, for `Array.prototype.sort` over many values: the order of two cells'
 * values, which it takes as `CellValue`s unchecked (`values.sort(valueComparator({ descending: true }))`).
 */
export const valueComparator = /* @__PURE__ */ strict(function valueComparator(opts?: { descending?: boolean }): (a: CellValue, b: CellValue) => number {
  const desc = descending(opts)
  return (a, b) => sortOrder(a, b, desc)
})
