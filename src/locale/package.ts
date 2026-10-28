/**
 * `loadLocale(tag)`: one of the package's locales loaded by its tag, its module imported as a thread of the pool
 * imports it (`modules.ts`) and loaded as `registerLocale` loads it, so a pool sends it to its threads. It needs
 * neither the engine nor the checks of `registerLocale`, the package's modules being its own: `xlcalc/xlsx` gives it too.
 */
import { InputError } from '../input-error.js'
import { checkString } from '../input/basic.js'
import { LOCALE_MODULES } from './modules.js'
import { addLocale, LOCALES, PACKAGE_LOCALES } from './registry.js'
import type { Locale } from './types.js'

/**
 * Loads one of the package's locales by its tag (`'ru-RU'`), as `import { ruRU } from 'xlcalc/locales/ru-RU'` and
 * `registerLocale(ruRU)` do: a workbook in that locale is taken from then on, and a pool's threads load it too. A locale
 * loaded already is given back without importing its module again; a tag that is not one of the package's locales is an
 * `InputError` (tags are compared as written).
 */
export async function loadLocale(tag: string): Promise<Locale> {
  const t = checkString(tag, 'tag')
  const had = LOCALES.get(t)
  if (had) return had
  const load = LOCALE_MODULES[t]
  if (!load) {
    const near = PACKAGE_LOCALES.find((x) => x.toLowerCase() === t.toLowerCase())
    throw new InputError('tag', `${JSON.stringify(t)} is not one of the package's locales (${PACKAGE_LOCALES.join(', ')})${near ? `; tags are case-sensitive: ${near}` : ''}`)
  }
  const locale = await load()
  // another call may have loaded it while the module was imported
  if (!LOCALES.has(t)) addLocale(locale)
  return LOCALES.get(t)!
}
