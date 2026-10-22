/**
 * Loading locales: `registerLocale`, which loads one of the package's locale modules for books in it, and the checks
 * of a locale given by the tag of a loaded one or as an object, whose shape is checked once.
 */
import { wordKey } from '../collate.js'
import { InputError } from '../input-error.js'
import { checkString, isObject, object, onlyKeys } from '../input/basic.js'
import { EN_MESSAGES } from '../locales/en-US.js'
import { XLSX_MESSAGE_KEYS } from './messages.js'
import { LOCALE_ERRORS } from '../syntax/errors.js'
import { TABLE_ITEMS } from '../syntax/table.js'
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
  if (!isObject(v)) throw new InputError(path, 'expected an object')
  const seen = new Map<string, string>()
  for (const [k, w] of Object.entries(v)) {
    if (keys && !keys.includes(k)) throw new InputError(path, `unknown key ${JSON.stringify(k)} (expected ${keys.join(', ')})`)
    if (typeof w !== 'string' || !rule.test(w)) throw new InputError(`${path}[${JSON.stringify(k)}]`, `expected ${expected}, got ${JSON.stringify(w)}`)
    const other = seen.get(wordKey(w))
    if (other !== undefined) throw new InputError(`${path}[${JSON.stringify(k)}]`, `${JSON.stringify(w)} is also the word for ${other}`)
    seen.set(wordKey(w), k)
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
  words(loc.errors, `${path}.errors`, LOCALE_ERRORS, /^#\S+$/u, 'an error literal starting with #')
  words(loc.functions, `${path}.functions`, null, WORD, 'a function name')
  if (loc.tableItems !== undefined) words(loc.tableItems, `${path}.tableItems`, TABLE_ITEMS, /^#[^[\]'@]+$/u, 'a specifier starting with #')
  if (loc.argumentWords !== undefined) argumentWords(loc.argumentWords, `${path}.argumentWords`)
  if (loc.functionText !== undefined) functionText(loc.functionText, `${path}.functionText`)
}

/** The words of arguments' names: two English words may share one (`criteria` and `condition` are `условие`). */
function argumentWords(v: unknown, path: string): void {
  if (!isObject(v)) throw new InputError(path, 'expected an object')
  for (const [k, w] of Object.entries(v))
    if (typeof w !== 'string' || !/^[\p{L}_][\p{L}\p{N}_]*$/u.test(w)) throw new InputError(`${path}[${JSON.stringify(k)}]`, `expected a word of an argument's name, got ${JSON.stringify(w)}`)
}

/** The words of the Insert Function window: texts by category and by function. */
function functionText(v: unknown, path: string): void {
  const o = object(v, path, 'an object such as { categories, descriptions }')
  for (const k of ['categories', 'descriptions'] as const) {
    const m = object(o[k], `${path}.${k}`, 'an object of texts')
    for (const [name, text] of Object.entries(m)) if (typeof text !== 'string' || !text) throw new InputError(`${path}.${k}[${JSON.stringify(name)}]`, `expected text, got ${JSON.stringify(text)}`)
  }
}

const checked = new WeakSet<object>()

/** The fields of a locale's rules of computation, every one of them required, and those it may leave out (en-US's then). */
const REGIONAL_KEYS: readonly (keyof RegionalData)[] = ['decimal', 'group', 'dateOrder', 'currency', 'months', 'monthsShort', 'days', 'daysShort', 'codes', 'colors', 'dateSeparator', 'lcid', 'longDate', 'longTime', 'longMonths', 'collation', 'builtin', 'codePage', 'columnWord']
const REGIONAL_OPTIONAL: readonly (keyof RegionalData)[] = ['formats']

/** A locale's rules of computation, where it has them: every field given, no other (en-US's where it has none). */
function checkRegional(v: unknown, path: string): void {
  const regional = object(v, path, "the rules a workbook is calculated by, as the package's locale modules have them")
  onlyKeys(regional, path, [...REGIONAL_KEYS, ...REGIONAL_OPTIONAL])
  for (const k of REGIONAL_KEYS) if (regional[k] === undefined) throw new InputError(`${path}.${k}`, 'expected a value')
}

/** A locale object, its formula text and rules checked once (a locale must not change after its first use). */
function checkLocaleObject(locale: unknown, path: string): Locale {
  if (typeof locale !== 'object' || locale === null || !checked.has(locale)) {
    checkText(locale, path)
    const regional = (locale as { regional?: unknown }).regional
    if (regional !== undefined) checkRegional(regional, `${path}.regional`)
    checked.add(locale as object)
  }
  return locale as Locale
}

/** A package's locale by name, as its module exports it (`ruRU`). */
const nameOf = (tag: string) => tag.replace('-', '')

/** How the root loads a package's locale: `import { ruRU } from 'xlcalc/locales/ru-RU'` and `registerLocale(ruRU)`. */
function howToLoad(tag: string): string {
  return `import { ${nameOf(tag)} } from 'xlcalc/locales/${tag}' and call registerLocale(${nameOf(tag)})`
}

/**
 * What a refusal of a package's locale not loaded says to do: without the root, which alone has `registerLocale`, give
 * the locale as an object; once the root is loaded (`rootLoaded`), import its module and load it. So the subpaths without
 * the engine (`./text`) carry neither the root's hint nor the package's name, which a host looks for as the engine's sign.
 */
let toLoad = (tag: string) => `pass the locale as an object: ${nameOf(tag)} from the package's module locales/${tag}`

/** The root is loaded: a package's locale not loaded is refused with its module to import and `registerLocale`. */
export function rootLoaded(): void {
  toLoad = howToLoad
}

/** A loaded locale's tag; a package's locale not loaded says how to load it (`hint`), another tag names the loaded ones (tags are compared as written). */
function checkTag(v: unknown, path: string, hint: (tag: string) => string): string {
  const tag = checkString(v, path)
  if (LOCALES.has(tag)) return tag
  const loaded = `loaded: ${[...LOCALES.keys()].join(', ')}`
  if (PACKAGE_LOCALES.includes(tag)) throw new InputError(path, `the locale ${tag} is not loaded: ${hint(tag)} (${loaded})`)
  const near = PACKAGE_LOCALES.find((t) => t.toLowerCase() === tag.toLowerCase())
  throw new InputError(path, `unknown locale ${JSON.stringify(tag)} (${loaded}${near ? `; tags are case-sensitive: ${near}` : `; the package's: ${PACKAGE_LOCALES.join(', ')}`})`)
}

/** A book's locale (`BookInput.locale`), the tag of a loaded one: a book takes no object, so a package's locale not loaded names its module and `registerLocale`, in the root, a thread or `xlsx`. */
export const checkBookLocale = (v: unknown, path: string): string => checkTag(v, path, howToLoad)

/** A locale given by the tag of a loaded one or as an object, loaded or not: the locale. */
export function checkLocaleArg(v: unknown, path: string): Locale {
  return typeof v === 'string' ? LOCALES.get(checkTag(v, path, toLoad))! : checkLocaleObject(v, path)
}

/**
 * Loads a locale for workbooks in that locale (`BookInput.locale`) and for its tag wherever a locale is accepted: one
 * of the package's locale modules (`import { ruRU } from 'xlcalc/locales/ru-RU'`, then `registerLocale(ruRU)`). Loading
 * the same object again does nothing; a different object with an already loaded tag throws an `InputError`. A pool's
 * threads load it too.
 */
export function registerLocale(locale: Locale): void {
  const l = checkLocaleObject(locale, 'locale')
  const had = LOCALES.get(l.tag)
  if (had === l) return
  if (had) throw new InputError('locale.tag', `the locale ${l.tag} is already loaded${l.tag === DEFAULT_LOCALE ? ': it is built in' : ', as another object'}`)
  // a loaded locale is a workbook's, which is calculated by its rules: they are required
  if (l.regional === undefined) checkRegional(l.regional, 'locale.regional')
  if (l.messages !== undefined) {
    const messages = object(l.messages, 'locale.messages', 'messages by key, each a function')
    // the engine's keys, whose English is en-US's, and the xlsx issues', whose English the subpath carries
    onlyKeys(messages, 'locale.messages', [...Object.keys(EN_MESSAGES), ...Object.keys(XLSX_MESSAGE_KEYS)])
    for (const [k, f] of Object.entries(messages)) if (typeof f !== 'function') throw new InputError(`locale.messages.${k}`, 'expected a function')
  }
  addLocale(l)
}
