/**
 * The loaded locales by BCP 47 tag: en-US, built in, and those `registerLocale` loaded (`load.ts`). A book's `locale`
 * names one of them; a locale object given where no book is needed works without loading.
 */
import { EN_US } from '../locales/en-US.js'
import type { Locale } from './types.js'

/** The locale of a book without one: Excel en-US. */
export const DEFAULT_LOCALE = 'en-US'

/**
 * The package's locales: en-US in the root, each other one the module `xlcalc/locales/<tag>`, which exports it by its tag
 * without the hyphen (`ruRU`). A thread of the pool imports the same modules (`worker/locales.ts`).
 */
export const PACKAGE_LOCALES: readonly string[] = ['en-US', 'ru-RU']

const loaded = new Map<string, Locale>([[DEFAULT_LOCALE, EN_US]])
const listeners = new Set<(locale: Locale) => void>()

/** The loaded locales by tag, en-US and the registered ones, in the order they were loaded. The map is for reading only. */
export const LOCALES: ReadonlyMap<string, Locale> = loaded

/** Loads a checked locale and tells the listeners (a pool sends it to its threads). */
export function addLocale(locale: Locale): void {
  loaded.set(locale.tag, locale)
  for (const f of listeners) f(locale)
}

/** Hears every locale loaded from now on; the answer stops it. */
export function onLocale(f: (locale: Locale) => void): () => void {
  listeners.add(f)
  return () => listeners.delete(f)
}

/** A locale object as it is, or the loaded locale of a tag; en-US without one, and for a tag not loaded (a book's tag is checked before). */
export const localeOf = (where: string | Locale | undefined): Locale => (typeof where === 'object' ? where : (loaded.get(where ?? DEFAULT_LOCALE) ?? EN_US))

/** The words of a book's locale for TRUE and FALSE, which are no names there (ИСТИНА and ЛОЖЬ in ru-RU). */
export function localeBooleans(where: string | Locale | undefined): string[] {
  const { booleans } = localeOf(where)
  return [booleans.TRUE, booleans.FALSE]
}
