/**
 * The locales of Excel's answers, by BCP 47 tag, in one place for the harnesses and the tests: what a command line names
 * (`en`, `ru`, `de-DE`), where a locale's answers are stored, how Excel for Mac starts in it and what Windows needs for
 * it. A locale added to `src/locales/` needs nothing here.
 */
import { registerLocale } from '../../src/locale/load.js'
import { DEFAULT_LOCALE, LOCALES } from '../../src/locale/registry.js'
import { displayRules } from '../../src/locale/rules.js'
import { LOCALE_MODULES } from '../../src/locale/modules.js'

// the harnesses compute in every locale of the package: its modules loaded, as a host loads the ones it needs
for (const load of Object.values(LOCALE_MODULES)) registerLocale(await load())

/** The package's locales, in the order of their tags; each has answers of Excel where they were taken. */
export const TAGS: readonly string[] = [...LOCALES.keys()].sort()

/** The short names a command line takes for en-US and ru-RU, which are also the names of their measurement files. */
const SHORT: Readonly<Record<string, string>> = { en: 'en-US', ru: 'ru-RU' }

/**
 * The locale a command line names: `en`, `ru` or a tag of a language and a region (`de-DE`); `fallback` when it names
 * none. The tag need not be one of the package's: Excel's answers can be taken before the locale's file is written.
 */
export function parseTag(arg: string | undefined, fallback: string = DEFAULT_LOCALE): string {
  if (arg === undefined) return fallback
  const tag = SHORT[arg] ?? arg
  if (!/^[a-z]{2,3}-[A-Z]{2}$/.test(tag)) throw new Error(`locale: en, ru or a tag of a language and a region such as de-DE, not ${JSON.stringify(arg)}`)
  return tag
}

/** The directory of a locale's answers within a kind's (`cases/`, `display/`, `tables/`): en-US's at the top, another's in one named by its tag. */
export const localeDir = (tag: string): string => (tag === DEFAULT_LOCALE ? '' : `${tag}/`)

/** The file of the coverage measurement's answers in a locale: `excel-en.json` and `excel-ru.json`, as they were first named, else `excel-<tag>.json`. */
export const measurementFile = (tag: string): string => `excel-${Object.keys(SHORT).find((k) => SHORT[k] === tag) ?? tag}.json`

/** The language and region of the Mac the harness runs on: Excel started without launch arguments computes in them. */
export const SYSTEM_TAG = 'ru-RU'

/**
 * Excel for Mac's launch arguments for a locale: none for the system's (ru-RU's answers were taken so), else its interface
 * language and region, as en-US's were taken (`-AppleLanguages (en) -AppleLocale en_US`). The interface language changes
 * computation, not only display: with the Russian one the text "TRUE" is no boolean (`IF("TRUE",1,2)` is #VALUE!).
 */
export function launchArgs(tag: string): string[] {
  if (tag === SYSTEM_TAG) return []
  const [language, region] = tag.split('-')
  return ['-AppleLanguages', `(${language})`, '-AppleLocale', `${language}_${region}`]
}

/** The locale an answer file's metadata names: the tag and how Excel was started. */
export const localeMeta = (tag: string): string => `${tag} (${launchArgs(tag).join(' ') || 'the system language and locale'})`

/**
 * What a locale needs on Windows: Excel's interface language id and the system code page of CHAR and CODE, from the
 * locale's data (`regional.lcid`, `regional.codePage`; en-US's are the engine's own, 1033 and 1252). Windows takes only
 * the package's locales.
 */
export function windowsLocale(tag: string): { lcid: number; codepage: number } {
  if (!LOCALES.has(tag)) throw new Error(`${tag} is none of the package's locales (${TAGS.join(', ')}): Windows takes its language id and code page from its file in src/locales/`)
  const lcid = tag === DEFAULT_LOCALE ? 0x409 : parseInt(LOCALES.get(tag)!.regional!.lcid, 16)
  return { lcid, codepage: Number(/^windows-(\d+)$/.exec(displayRules(tag).codePage.name)?.[1]) }
}
