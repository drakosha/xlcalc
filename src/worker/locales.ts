/**
 * The package's locale modules by tag, which a thread of the pool imports itself when the pool sends it one of their
 * locales: a locale's messages are functions, which no message between threads carries (`registry.ts` lists the tags).
 */
import type { Locale } from '../locale/types.js'

export const LOCALE_MODULES: Readonly<Record<string, () => Promise<Locale>>> = {
  'ru-RU': async () => (await import('../locales/ru-RU.js')).ruRU,
}
