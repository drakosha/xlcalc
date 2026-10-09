/**
 * The package's locales by BCP 47 tag, read from the list of locale files (`src/locales/index.ts`): a book's `locale`
 * names one of them, and its formula text and rules of computation come from its file.
 */
import * as LIST from '../locales/index.js'
import type { LocaleFile } from './file.js'
import type { Locale } from './types.js'

/** The locale of a book without one: Excel en-US. */
export const DEFAULT_LOCALE = 'en-US'

/** The locale files by tag, in the order of their tags. */
export const LOCALE_FILES: ReadonlyMap<string, LocaleFile> = new Map(
  Object.values(LIST)
    .map((f): [string, LocaleFile] => [f.text.tag, f])
    .sort(([a], [b]) => (a < b ? -1 : 1)),
)

/**
 * The package's locales by BCP 47 tag, each its formula text for `toLocal`, `toInvariant` and `validateLocalFormula`
 * (`LOCALES.get('ru-RU')`); a book's `locale` and `displayValue`'s take these tags. The map is for reading: the engine
 * checks tags against its own list.
 */
export const LOCALES: ReadonlyMap<string, Locale> = new Map([...LOCALE_FILES].map(([tag, f]) => [tag, f.text]))

/** The words of a book's locale for TRUE and FALSE, which are no names there (ИСТИНА and ЛОЖЬ in ru-RU). */
export function localeBooleans(tag: string | undefined): string[] {
  const loc = LOCALE_FILES.get(tag ?? DEFAULT_LOCALE)?.text
  return loc ? [loc.booleans.TRUE, loc.booleans.FALSE] : []
}
