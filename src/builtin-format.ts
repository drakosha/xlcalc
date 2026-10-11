/**
 * The code of one of Excel's built-in number formats as Excel shows it in a locale (`builtinFormatCode`): a cell's
 * `format: 14` is the short date, `m/d/yyyy` in en-US and `dd.mm.yyyy` in ru-RU. The codes are the tables the engine
 * shows the formats by (`numfmt/builtin.ts`, a locale's `builtin`, through `builtinIn`), in en-US syntax as a cell's
 * `format` takes a code.
 */
import { checkLocale, InputError, what } from './input.js'
import { builtinIn } from './locale/rules.js'
import { isBuiltin } from './numfmt.js'

/**
 * The code a built-in number format (0 to 49, `NumberFormat` as a number) shows by in `locale` (en-US without one), in
 * en-US syntax as a cell's `format` takes it; undefined for a number that is not a built-in format's.
 */
export function builtinFormatCode(id: number, locale?: string): string | undefined {
  if (typeof id !== 'number') throw new InputError('id', `expected the number of a built-in format, 0 to 49, got ${what(id)}`)
  if (locale !== undefined) checkLocale(locale, 'locale')
  return isBuiltin(id) ? builtinIn(locale)(id) : undefined
}
