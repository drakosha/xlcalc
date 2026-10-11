/**
 * Loading locales: `registerLocale`, which loads one of the package's locale modules for books in it, and the checks
 * of a locale given by the tag of a loaded one or as an object, whose shape is checked once.
 */
import { checkString, InputError, object, onlyKeys } from '../input.js'
import { EN_MESSAGES } from '../locales/en-US.js'
import { ERROR_LITERALS } from './lex.js'
import { addLocale, DEFAULT_LOCALE, LOCALES, PACKAGE_LOCALES } from './registry.js'
import type { RegionalData } from './rules/data.js'
import type { Locale } from './types.js'

const WORD = /^[\p{L}_][\p{L}\p{N}_.]*$/u
// characters that start or end other pieces of a formula, or that a name holds (`?`); `:` and `\` only inside array
// constants (ru-RU: `{1;2:3;4}`)
const RESERVED = /[\p{L}\p{N}\s"'#(){}[\]!$+\-*/^&=<>%_@?]/u

function separator(loc: Record<string, unknown>, path: string, key: string, inArrays: boolean): string {
  const v = checkString(loc[key], `${path}.${key}`)
  if (v.length !== 1 || RESERVED.test(v) || (!inArrays && (v === ':' || v === '\\'))) throw new InputError(`${path}.${key}`, `expected one character that has no other meaning in a formula, got ${JSON.stringify(v)}`)
  return v
}

function words(v: unknown, path: string, keys: readonly string[] | null, rule: RegExp, expected: string): void {
  if (typeof v !== 'object' || v === null || Array.isArray(v)) throw new InputError(path, 'expected an object')
  const seen = new Map<string, string>()
  for (const [k, w] of Object.entries(v)) {
    if (keys && !keys.includes(k)) throw new InputError(path, `unknown key ${JSON.stringify(k)} (expected ${keys.join(', ')})`)
    if (typeof w !== 'string' || !rule.test(w)) throw new InputError(`${path}[${JSON.stringify(k)}]`, `expected ${expected}, got ${JSON.stringify(w)}`)
    const other = seen.get(w.toUpperCase())
    if (other !== undefined) throw new InputError(`${path}[${JSON.stringify(k)}]`, `${JSON.stringify(w)} is also the word for ${other}`)
    seen.set(w.toUpperCase(), k)
  }
}

/** The shape of a locale's formula text: separators that cannot be confused, words that read as words. */
function checkText(locale: unknown, path: string): void {
  const loc = object(locale, path, "a locale's tag or a locale object such as ruRU")
  checkString(loc.tag, `${path}.tag`)
  const decimal = separator(loc, path, 'decimalSeparator', false)
  const list = separator(loc, path, 'listSeparator', false)
  const col = separator(loc, path, 'arrayColumnSeparator', true)
  const row = separator(loc, path, 'arrayRowSeparator', true)
  if (decimal === list || decimal === col || decimal === row || col === row) throw new InputError(path, 'the decimal separator, the list separator and the array separators must differ (the list separator may equal an array separator)')
  words(loc.booleans, `${path}.booleans`, ['TRUE', 'FALSE'], WORD, 'a word')
  if (Object.keys(loc.booleans as object).length !== 2) throw new InputError(`${path}.booleans`, 'expected both TRUE and FALSE')
  words(loc.errors, `${path}.errors`, ERROR_LITERALS, /^#\S+$/u, 'an error literal starting with #')
  words(loc.functions, `${path}.functions`, null, WORD, 'a function name')
  if (loc.tableItems !== undefined) words(loc.tableItems, `${path}.tableItems`, ['#All', '#Data', '#Headers', '#Totals', '#This Row'], /^#[^[\]'@]+$/u, 'a specifier starting with #')
}

const checked = new WeakSet<object>()

/** A locale object, its formula text checked once (a locale must not change after its first use). */
function checkLocaleObject(locale: unknown, path: string): Locale {
  if (typeof locale !== 'object' || locale === null || !checked.has(locale)) {
    checkText(locale, path)
    checked.add(locale as object)
  }
  return locale as Locale
}

/** How a package's locale is loaded: `import { ruRU } from 'xlcalc/locales/ru-RU'` and `registerLocale(ruRU)`. */
function howToLoad(tag: string): string {
  const name = tag.replace('-', '')
  return `import { ${name} } from 'xlcalc/locales/${tag}' and call registerLocale(${name})`
}

/** A loaded locale's tag; a package's locale not loaded names its module, another tag the loaded ones (tags are compared as written). */
export function checkLocaleTag(v: unknown, path: string): string {
  const tag = checkString(v, path)
  if (LOCALES.has(tag)) return tag
  const loaded = `loaded: ${[...LOCALES.keys()].join(', ')}`
  if (PACKAGE_LOCALES.includes(tag)) throw new InputError(path, `the locale ${tag} is not loaded: ${howToLoad(tag)} (${loaded})`)
  const near = PACKAGE_LOCALES.find((t) => t.toLowerCase() === tag.toLowerCase())
  throw new InputError(path, `unknown locale ${JSON.stringify(tag)} (${loaded}${near ? `; tags are case-sensitive: ${near}` : `; the package's: ${PACKAGE_LOCALES.join(', ')}`})`)
}

/** A locale given by the tag of a loaded one or as an object, loaded or not: the locale. */
export function checkLocaleArg(v: unknown, path: string): Locale {
  return typeof v === 'string' ? LOCALES.get(checkLocaleTag(v, path))! : checkLocaleObject(v, path)
}

/** The fields of a locale's rules of computation, every one of them required. */
const REGIONAL_KEYS: readonly (keyof RegionalData)[] = ['decimal', 'group', 'dateOrder', 'currency', 'months', 'monthsShort', 'days', 'daysShort', 'codes', 'colors', 'dateSeparator', 'lcid', 'longDate', 'longTime', 'longMonths', 'collation', 'builtin', 'codePage', 'columnWord']

/**
 * Loads a locale for books in it (`BookInput.locale`) and for its tag wherever a locale is taken: one of the package's
 * locale modules (`import { ruRU } from 'xlcalc/locales/ru-RU'`, then `registerLocale(ruRU)`). Loading the same object
 * again does nothing; another object with a loaded tag is an `InputError`. A pool's threads load it too.
 */
export function registerLocale(locale: Locale): void {
  const l = checkLocaleObject(locale, 'locale')
  const had = LOCALES.get(l.tag)
  if (had === l) return
  if (had) throw new InputError('locale.tag', `the locale ${l.tag} is loaded already${l.tag === DEFAULT_LOCALE ? ': it is built in' : ', as another object'}`)
  const regional = object(l.regional, 'locale.regional', "the rules a book computes by, as the package's locale modules have them")
  onlyKeys(regional, 'locale.regional', REGIONAL_KEYS)
  for (const k of REGIONAL_KEYS) if (regional[k] === undefined) throw new InputError(`locale.regional.${k}`, 'expected a value')
  if (l.messages !== undefined) {
    const messages = object(l.messages, 'locale.messages', 'messages by key, each a function')
    onlyKeys(messages, 'locale.messages', Object.keys(EN_MESSAGES))
    for (const [k, f] of Object.entries(messages)) if (typeof f !== 'function') throw new InputError(`locale.messages.${k}`, 'expected a function')
  }
  addLocale(l)
}
