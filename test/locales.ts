/**
 * The package's locale modules, loaded for the tests that compute in them: importing this file registers each one, as
 * a host does (`registerLocale`). The tests of a locale not loaded do not import it, and the tests of en-US need it not.
 */
import { registerLocale, type Locale } from '../src/index.js'
import { LOCALE_MODULES } from '../src/locale/modules.js'

export { ruRU } from '../src/locales/ru-RU.js'

/** Every locale module of the package (`src/locales/`, but en-US, which is built in). */
export const MODULES: readonly Locale[] = await Promise.all(Object.values(LOCALE_MODULES).map((load) => load()))

for (const l of MODULES) registerLocale(l)
