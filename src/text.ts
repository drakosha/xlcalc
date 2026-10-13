/**
 * Formula text without the engine (the `./text` subpath, the same functions in the root): the references of text as it
 * is typed, for colouring them, a sheet's name and a table column's name as a reference writes them, the clipboard's
 * text as the rectangle Excel pastes, and the category of a number format as the ribbon's Number Format box shows it,
 * with the items of its list (`numfmt/category.ts`, by the engine's reader of format codes). The tokenizer and the
 * translation of a locale's formula text are the engine's own (`scan.ts`, `locale/translate.ts`); they compute nothing,
 * and no locale's data comes with them but en-US's messages, which the tokenizer's errors carry: another locale is given
 * as an object (`ruRU` of `xlcalc/locales/ru-RU`). Each function refuses options it does not take (`strict-args.ts`).
 */
import { checkString } from './input.js'
import { clipboardRows } from './entry/clipboard.js'
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
