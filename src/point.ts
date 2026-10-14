/**
 * What an arrow does in formula text typed over a cell (Excel's Enter mode), measured on Excel for Mac in en-US and ru-RU
 * (`bench/coverage/ui/point-*.ts`, `docs/entry.md`, "Point mode"): text that does not begin a formula is entered; in a
 * formula a reference goes in where an operand is wanted (after the formula's first character, an operator, `(`, the
 * list separator, `:`, a sheet's `!`, a space after an operand), the formula is entered after an operand (a number, a
 * reference, a name, a closing bracket, a closed string), and nothing happens inside a string, a quoted sheet's name or
 * a table's brackets, or after a character the locale's formulas do not have. A host that pointed already moves its own
 * reference: this says what an arrow does with no reference pointed since the text was typed.
 */
import type { PointAction, ScanOptions } from './contract.js'
import { checkNumber, checkString, object, onlyKeys } from './input.js'
import { checkLocaleArg } from './locale/load.js'
import { isSpace } from './syntax/scan.js'

/** The characters after which an operand is wanted, beside the list separator: the operators, `(`, the range's `:`, a sheet's `!`. */
const WANTS_OPERAND = new Set([...'=+-*/^&<>(:!@'])

/** The last character of an operand: a letter or a digit (a number, a reference, a name), a closing bracket, a closed string, `%`. */
const ENDS_OPERAND = /[\p{L}\p{N}_)"}%#]/u

/** Whether digits end at `end` of `text` as a number of their own: not the last of a reference's or a name's characters. */
function numberEndsAt(text: string, end: number): boolean {
  let j = end
  while (j > 0 && /\p{N}/u.test(text[j - 1])) j--
  return j < end && !/[\p{L}_$.\\]/u.test(text[j - 1] ?? '')
}

/** Whether `upTo` lies inside a string, a quoted sheet's name or a table's brackets of `text`. */
function inLiteral(text: string, upTo: number): boolean {
  let quote: string | null = null
  let depth = 0
  for (let i = 0; i < upTo; i++) {
    const c = text[i]
    if (quote) {
      if (c === quote) quote = null
    } else if (c === '"' || c === "'") quote = c
    else if (c === '[') depth++
    else if (c === ']') depth = Math.max(0, depth - 1)
  }
  return quote !== null || depth > 0
}

/**
 * What an arrow does at `caret` (the end of the text by default) of `text` typed over a cell: a reference put in at the
 * caret, the text entered, or nothing (`PointAction`). `opts.locale`: the locale the text is typed in (its list and
 * decimal separators); without it en-US's.
 */
export function refAt(text: string, caret?: number, opts?: ScanOptions): PointAction {
  const t = checkString(text, 'text')
  if (caret !== undefined) checkNumber(caret, 'caret', 0, t.length, true)
  const at = caret ?? t.length
  if (opts !== undefined) onlyKeys(object(opts, 'opts', "options such as { locale: 'ru-RU' }"), 'opts', ['locale'])
  const locale = opts?.locale === undefined ? undefined : checkLocaleArg(opts.locale, 'opts.locale')
  const list = locale?.listSeparator ?? ','
  const decimal = locale?.decimalSeparator ?? '.'
  // a value: the arrow enters it, as Excel does any text not begun by `=`, `+` or `-`
  if (!/^[=+-]/.test(t)) return { action: 'commit' }
  if (inLiteral(t, at)) return { action: 'none' }
  let i = at
  while (i > 1 && isSpace(t[i - 1])) i--
  // the formula's first character alone
  if (i <= 1) return { action: 'insert', at }
  const ch = t[i - 1]
  if (WANTS_OPERAND.has(ch) || ch === list) return { action: 'insert', at }
  if (ENDS_OPERAND.test(ch)) return i < at ? { action: 'insert', at } : { action: 'commit' }
  // a number's point with no digit after it (`=1.`) takes a reference, as Excel's measured answer has it; digits that end
  // a reference or a name (`A1,` in ru-RU) are no number
  if ((ch === '.' || ch === decimal) && numberEndsAt(t, i - 1)) return { action: 'insert', at }
  if (ch === '.') return { action: 'commit' }
  return { action: 'none' }
}
