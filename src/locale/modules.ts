/**
 * The package's locale modules by tag, which `loadLocale` imports, and a thread of the pool when the pool sends it one
 * of their locales: a locale's messages are functions, which no message between threads carries (`registry.ts` lists
 * the tags).
 */
import type { Locale } from './types.js'

export const LOCALE_MODULES: Readonly<Record<string, () => Promise<Locale>>> = {
  'ru-RU': async () => (await import('../locales/ru-RU.js')).ruRU,
}
