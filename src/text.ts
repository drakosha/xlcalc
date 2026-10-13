/**
 * Formula text without the engine (the `./text` subpath, the same functions in the root): the references of text as it
 * is typed, for colouring them, a sheet's name and a table column's name as a reference writes them, the clipboard's
 * text as the rectangle Excel pastes, and the category of a number format as the ribbon's Number Format box shows it,
 * with the items of its list (`numfmt/category.ts`, by the engine's reader of format codes). The tokenizer and the
 * translation of a locale's formula text are the engine's own (`scan.ts`, `locale/translate.ts`); they compute nothing,
 * and no locale's data comes with them but en-US's messages, which the tokenizer's errors carry: another locale is given
 * as an object (`ruRU` of `xlcalc/locales/ru-RU`). Each function refuses options it does not take (`strict-args.ts`).
 */
import { array, checkBoolean, checkNumber, checkString, InputError, object, onlyOptions } from './input.js'
import { clipboardRows } from './entry/clipboard.js'
import { columnPoints, widthOf } from './column-width.js'
import * as category from './numfmt/category.js'
import * as scan from './scan.js'
import { strict } from './strict-args.js'
import { escapeColumn as escaped } from './syntax/table.js'

export { quoteSheetIfNeeded } from './a1.js'

export const scanReferences = /* @__PURE__ */ strict(scan.scanReferences)
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

/** Excel's plain-text copy: tabs, CRLF rows and quoted control characters; Windows ends with CRLF, Mac does not. */
export const formatClipboard = /* @__PURE__ */ strict(function formatClipboard(rows: string[][], opts?: { trailingNewline?: boolean }): string {
  const o = opts === undefined ? {} : object(opts, 'opts', 'clipboard options such as { trailingNewline: false }')
  onlyOptions(o, 'opts', ['trailingNewline'])
  if (o.trailingNewline !== undefined) checkBoolean(o.trailingNewline, 'opts.trailingNewline')
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
