/**
 * What a locale file of `src/locales/` holds: everything the engine knows of the locale, as data. The package's locales
 * are the files listed in `src/locales/index.ts`, one line each (`registry.ts` reads the list).
 */
import type { Messages } from './messages.js'
import type { RegionalData } from './rules/data.js'
import type { Locale } from './types.js'

export interface LocaleFile {
  /** Formula text: separators, the words for TRUE and FALSE, error literals, function names (`toLocal`, `toInvariant`). */
  readonly text: Locale
  /** The rules of computation as data (`regionalRules`); en-US has none, its rules are the engine's own code (`rules.ts`). */
  readonly regional?: RegionalData
  /**
   * Widths of the letters of the locale's script in Calibri 11 in Excel for Mac, in points, by characters of equal width,
   * for what a cell shows by its column's width (`numfmt/glyphs.ts`, where Latin letters, digits and signs are); they
   * count in books of every locale.
   */
  readonly glyphWidths?: readonly (readonly [number, string])[]
  /**
   * Messages for people in the locale's language (`messages.ts`): why a formula does not read, why a structure operation
   * is refused, the words of `explain`. en-US has every one; a message the locale lacks is en-US's.
   */
  readonly messages?: Partial<Messages>
}
