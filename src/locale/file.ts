/**
 * What a locale file of `src/locales/` holds: everything the engine knows of the locale, as data. The package's locales
 * are the files listed in `src/locales/index.ts`, one line each (`registry.ts` reads the list).
 */
import type { RegionalData } from './rules/data.js'
import type { Locale } from './types.js'

export interface LocaleFile {
  /** Formula text: separators, the words for TRUE and FALSE, error literals, function names (`toLocal`, `toInvariant`). */
  readonly text: Locale
  /** The rules of computation as data (`regionalRules`); en-US has none, its rules are the engine's own code (`rules.ts`). */
  readonly regional?: RegionalData
}
