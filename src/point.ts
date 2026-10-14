/**
 * What an arrow does in formula text typed over a cell (Excel's Enter mode), measured on Excel for Mac in en-US and ru-RU
 * (`bench/coverage/ui/point-*.ts`, `docs/entry.md`, "Point mode"): text that does not begin a formula is entered; in a
 * formula a reference goes in where an operand is wanted (after the formula's first character, an operator, `(`, the
 * list separator, `:`, a sheet's `!`, a point, a space), and elsewhere the formula is entered when it reads as it stands,
 * in the locale it is typed in, and nothing happens when it does not (`=A1)`, `=1e`, `=1.5` in ru-RU) or when the caret
 * is inside a string, a quoted sheet's name or a table's brackets. A host that pointed already moves its own reference:
 * this says what an arrow does with no reference pointed since the text was typed.
 */
import type { PointAction, ScanOptions } from './contract.js'
import { checkNumber, checkString, object, onlyKeys } from './input.js'
import { sidesOf } from './locale/index.js'
import { checkLocaleArg } from './locale/load.js'
import { ownErrors, translate } from './locale/translate.js'
import type { Locale } from './locale/types.js'
import { FormulaError } from './syntax/ast.js'
import { parse } from './syntax/parse.js'
import { isSpace } from './syntax/scan.js'

/** The characters after which an operand is wanted, beside the list separator: the operators, `(`, the range's `:`, a sheet's `!`, and a point (`=1.`, `=A1.`, as Excel's answers have it). */
const WANTS_OPERAND = new Set([...'=+-*/^&<>(:!@.'])

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

/** Whether a formula's text after its first character reads as a formula, in `locale` (en-US's stored syntax without one). */
function reads(body: string, locale: Locale | undefined): boolean {
  let stored = body
  if (locale) {
    const { local, invariant } = sidesOf(locale)
    const t = translate(body, ownErrors(locale, local), invariant)
    if (!t.ok) return false
    stored = t.formula
  }
  try {
    parse(stored)
    return true
  } catch (e) {
    if (e instanceof FormulaError) return false
    throw e
  }
}

/**
 * What an arrow does at `caret` (the end of the text by default) of `text` typed over a cell: a reference put in at the
 * caret, the text entered, or nothing (`PointAction`). `opts.locale`: the locale the text is typed in, a loaded locale's
 * tag or a locale object; without it the text is en-US's.
 */
export function refAt(text: string, caret?: number, opts?: ScanOptions): PointAction {
  const t = checkString(text, 'text')
  if (caret !== undefined) checkNumber(caret, 'caret', 0, t.length, true)
  const at = caret ?? t.length
  if (opts !== undefined) onlyKeys(object(opts, 'opts', "options such as { locale: 'ru-RU' }"), 'opts', ['locale'])
  const locale = opts?.locale === undefined ? undefined : checkLocaleArg(opts.locale, 'opts.locale')
  // a value: the arrow enters it, as Excel does any text not begun by `=`, `+` or `-`
  if (!/^[=+-]/.test(t)) return { action: 'commit' }
  if (inLiteral(t, at)) return { action: 'none' }
  let i = at
  while (i > 1 && isSpace(t[i - 1])) i--
  // after the formula's first character, an operand wanted, or a space after an operand (the intersection operator);
  // before the first character no reference goes in
  if (at > 0 && (i <= 1 || i < at || WANTS_OPERAND.has(t[i - 1]) || t[i - 1] === (locale?.listSeparator ?? ','))) return { action: 'insert', at }
  // the arrow enters the whole text, wherever the caret is: it goes in when the formula reads
  return reads(t.startsWith('=') ? t.slice(1) : t, locale) ? { action: 'commit' } : { action: 'none' }
}
