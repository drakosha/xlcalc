/**
 * The package's locale modules by tag, which a thread of the pool imports itself when the pool sends it one of their
 * locales: a locale's messages are functions, which no message between threads carries (`registry.ts` lists the tags);
 * and the module of match mode 3's levels, whose registered one a thread imports by its kind.
 */
import type { RegexLevel } from '../contract.js'
import type { Locale } from '../locale/types.js'

export const LOCALE_MODULES: Readonly<Record<string, () => Promise<Locale>>> = {
  'ru-RU': async () => (await import('../locales/ru-RU.js')).ruRU,
}

export const REGEX_MODULE = async (kind: RegexLevel['kind']): Promise<RegexLevel> => {
  const m = await import('../regex/index.js')
  return kind === 'xlcalc-pcre-syntax' ? m.pcreSyntax : m.pcreEngine
}
