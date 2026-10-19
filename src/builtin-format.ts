/**
 * The code of one of Excel's built-in number formats as Excel shows it in a locale (`builtinFormatCode`): a cell's
 * `numberFormat: 14` is the short date, `m/d/yyyy` in en-US and `dd.mm.yyyy` in ru-RU. The codes are the tables the engine
 * shows the formats by (`numfmt/builtin.ts`, a locale's `builtin`, through `builtinIn`), in en-US syntax as a cell's
 * `numberFormat` takes a code.
 */
import { InputError, what } from './input-error.js'
import { checkLocaleArg } from './locale/load.js'
import { builtinIn } from './locale/rules.js'
import type { Locale } from './locale/types.js'
import { isBuiltin } from './numfmt.js'

/**
 * The format code with which a built-in number format (0 to 49, `NumberFormat` as a number) is displayed in `locale` (a
 * loaded locale's tag or a locale object; en-US without one), in en-US syntax as a cell's `numberFormat` accepts it;
 * undefined for a number that is not a built-in format.
 */
export function builtinFormatCode(id: number, locale?: string | Locale): string | undefined {
  if (typeof id !== 'number') throw new InputError('id', `expected the number of a built-in format, 0 to 49, got ${what(id)}`)
  const loc = locale === undefined ? undefined : checkLocaleArg(locale, 'locale')
  return isBuiltin(id) ? builtinIn(loc)(id) : undefined
}
