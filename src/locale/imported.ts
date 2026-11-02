/**
 * Where a package's locale module puts its locale on being imported (`importedLocale`), without the registry and its
 * en-US: the registry takes them as it starts, and each one imported after (`registry.ts`).
 */
import type { Locale } from './types.js'

/** The locales the package's modules loaded on being imported, which `registerLocale` may replace with a host's own copy. */
export const imported = new Set<Locale>()
let sink: ((locale: Locale) => void) | undefined

/** Loads a package module's locale as the module is imported: importing it is enough. */
export function importedLocale(locale: Locale): void {
  imported.add(locale)
  sink?.(locale)
}

/** The registry's: takes the locales imported so far and each one imported from now on. */
export function takeImported(f: (locale: Locale) => void): void {
  sink = f
  imported.forEach(f)
}
