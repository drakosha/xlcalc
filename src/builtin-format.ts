/**
 * The code of one of Excel's built-in number formats as Excel shows it in a locale (`builtinFormatCode`): a cell's
 * `format: 14` is the short date, `m/d/yyyy` in en-US and `dd.mm.yyyy` in ru-RU. The codes are the tables the engine
 * shows the formats by (`numfmt/builtin.ts`, a locale's `builtin`), as an xlsx file writes a code (en-US syntax).
 */
import { checkLocale, InputError, what } from './input.js'
import { DEFAULT_LOCALE } from './locale/registry.js'
import { REGIONAL } from './locale/rules.js'
import { builtinCodes, EN_BUILTIN, isBuiltin } from './numfmt.js'

/** The codes a locale shows the built-in formats by, by number: its own, en-US's where it has none. */
export function builtinCodesOf(tag: string): (id: number) => string {
  const own = REGIONAL.get(tag)?.builtin
  return own ? builtinCodes(own) : (id) => EN_BUILTIN[id]
}

/**
 * The code a built-in number format (0 to 49, `NumberFormat` as a number) shows by in `locale` (en-US without one), in
 * en-US syntax as a cell's `format` takes it; undefined for a number that is not a built-in format's.
 */
export function builtinFormatCode(id: number, locale?: string): string | undefined {
  if (typeof id !== 'number') throw new InputError('id', `expected the number of a built-in format, 0 to 49, got ${what(id)}`)
  if (locale !== undefined) checkLocale(locale, 'locale')
  return isBuiltin(id) ? builtinCodesOf(locale ?? DEFAULT_LOCALE)(id) : undefined
}
