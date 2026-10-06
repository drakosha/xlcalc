/** The package's locales by BCP 47 tag: a book's `locale` names one of them (one line per locale file). */
import { enUS } from './en-US.js'
import { ruRU } from './ru-RU.js'
import type { Locale } from './types.js'

/** The locale of a book without one: Excel en-US. */
export const DEFAULT_LOCALE = 'en-US'

export const LOCALES: ReadonlyMap<string, Locale> = new Map([
  ['en-US', enUS],
  ['ru-RU', ruRU],
])
