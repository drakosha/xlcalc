/**
 * Locales of formula text: typing and showing formulas the way Excel does in another language, while books
 * keep the invariant text. `toLocal` shows a stored formula, `toInvariant` reads typed text back, and
 * `validateLocalFormula` checks typed text with positions in it.
 */
import type { ValidateContext, ValidateResult } from '../contract.js'
import { checkString, checkValidateContext, InputError } from '../input.js'
import { validate } from '../syntax.js'
import { ERROR_LITERALS } from './lex.js'
import { sides, translate, type Sides } from './translate.js'
import type { Locale, Translation } from './types.js'

export { enUS } from './en-US.js'
export { ruRU } from './ru-RU.js'
export type { Locale, Translation } from './types.js'

const WORD = /^[\p{L}_][\p{L}\p{N}_.]*$/u
// characters that start or end other pieces of a formula; `:` and `\` only inside array constants (ru-RU: `{1;2:3;4}`)
const RESERVED = /[\p{L}\p{N}\s"'#(){}[\]!$+\-*/^&=<>%_@]/u

function separator(loc: Record<string, unknown>, key: string, inArrays: boolean): string {
  const v = checkString(loc[key], `locale.${key}`)
  if (v.length !== 1 || RESERVED.test(v) || (!inArrays && (v === ':' || v === '\\'))) throw new InputError(`locale.${key}`, `expected one character that has no other meaning in a formula, got ${JSON.stringify(v)}`)
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

/** The shape of a locale: separators that cannot be confused, words that read as words. */
function checkLocale(locale: unknown): Locale {
  if (typeof locale !== 'object' || locale === null || Array.isArray(locale)) throw new InputError('locale', 'expected a locale object such as ruRU')
  const loc = locale as Record<string, unknown>
  checkString(loc.tag, 'locale.tag')
  const decimal = separator(loc, 'decimalSeparator', false)
  const list = separator(loc, 'listSeparator', false)
  const col = separator(loc, 'arrayColumnSeparator', true)
  const row = separator(loc, 'arrayRowSeparator', true)
  if (decimal === list || decimal === col || decimal === row || col === row) throw new InputError('locale', 'the decimal separator, the list separator and the array separators must differ (the list separator may equal an array separator)')
  words(loc.booleans, 'locale.booleans', ['TRUE', 'FALSE'], WORD, 'a word')
  if (Object.keys(loc.booleans as object).length !== 2) throw new InputError('locale.booleans', 'expected both TRUE and FALSE')
  words(loc.errors, 'locale.errors', ERROR_LITERALS, /^#\S+$/u, 'an error literal starting with #')
  words(loc.functions, 'locale.functions', null, WORD, 'a function name')
  return locale as Locale
}

const cache = new WeakMap<object, Sides>()

/** A locale's translation tables, built once per locale object (a locale must not change after its first use). */
function sidesOf(locale: Locale): Sides {
  let s = cache.get(locale)
  if (!s) cache.set(locale, (s = sides(checkLocale(locale))))
  return s
}

const result = (r: ReturnType<typeof translate>): Translation => (r.ok ? { ok: true, formula: r.formula } : r)

/** A stored formula as `locale` shows it: `SUM(A1:A3,1.5)` is `СУММ(A1:A3;1,5)` in ru-RU. */
export function toLocal(formula: string, locale: Locale): Translation {
  const s = sidesOf(locale)
  return result(translate(checkString(formula, 'formula'), s.invariant, s.local))
}

/**
 * Text typed in `locale` as a stored formula: `=СУММ(A1:A3;1,5)` is `SUM(A1:A3,1.5)`. A leading `=` is
 * dropped, as stored formulas have none; English function names, TRUE/FALSE and error texts are read too.
 * Positions of errors are in `text`.
 */
export function toInvariant(text: string, locale: Locale): Translation {
  const s = sidesOf(locale)
  const eq = checkString(text, 'text').startsWith('=') ? 1 : 0
  return result(translate(text.slice(eq), s.local, s.invariant, eq))
}

/**
 * `validateFormula` for text typed in `locale`: the same result, with the canonical stored text, and
 * positions of errors in `text` (a leading `=` is allowed).
 */
export function validateLocalFormula(text: string, locale: Locale, ctx?: ValidateContext): ValidateResult {
  const s = sidesOf(locale)
  const eq = checkString(text, 'text').startsWith('=') ? 1 : 0
  const c = checkValidateContext(ctx)
  const t = translate(text.slice(eq), s.local, s.invariant, eq)
  if (!t.ok) return t
  const r = validate(t.formula, c)
  if (!r.ok) return { ...r, at: t.back(r.at) }
  const { ast: _ast, ...rest } = r
  return rest
}
