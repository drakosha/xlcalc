/**
 * Locales of formula text: typing and showing formulas the way Excel does in another language, while books
 * keep the invariant text. `toLocal` shows a stored formula, `toInvariant` reads typed text back, and
 * `validateLocalFormula` checks typed text with positions in it.
 */
import type { ValidateContext, ValidateResult } from '../contract.js'
import { checkString, checkValidateContext, InputError } from '../input.js'
import { FormulaError, parse, sheetKey, tokenize, validate, type Tok } from '../syntax.js'
import { shownSpec, unnamedInTable } from '../syntax/table.js'
import { EN_US } from '../locales/en-US.js'
import { RU_RU } from '../locales/ru-RU.js'
import { ERROR_LITERALS } from './lex.js'
import { sides, translate, type Sides } from './translate.js'
import type { Locale, Translation } from './types.js'

export type { Locale, Translation } from './types.js'

/** en-US's and ru-RU's formula text by name; every locale's is in `LOCALES` by tag. */
export const enUS: Locale = EN_US.text
export const ruRU: Locale = RU_RU.text

const WORD = /^[\p{L}_][\p{L}\p{N}_.]*$/u
// characters that start or end other pieces of a formula, or that a name holds (`?`); `:` and `\` only inside array
// constants (ru-RU: `{1;2:3;4}`)
const RESERVED = /[\p{L}\p{N}\s"'#(){}[\]!$+\-*/^&=<>%_@?]/u

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
  if (loc.tableItems !== undefined) words(loc.tableItems, 'locale.tableItems', ['#All', '#Data', '#Headers', '#Totals', '#This Row'], /^#[^[\]'@]+$/u, 'a specifier starting with #')
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

/**
 * A stored formula as `locale` shows it: `SUM(A1:A3,1.5)` is `СУММ(A1:A3;1,5)` in ru-RU. Structured references are shown
 * as Excel shows them, the formula's row as `@` (`Sales[[#This Row],[Qty]]` is `Sales[@Qty]`) and a table alone by its
 * name (`Sales[]` is `Sales`); `opts.table`, the table of the formula's cell, drops its name before a column of its own
 * (`[@Qty]`, `SUM([Qty])`), as Excel shows a formula inside a table.
 */
export function toLocal(formula: string, locale: Locale, opts: { table?: string } = {}): Translation {
  const s = sidesOf(locale)
  if (opts.table !== undefined) checkString(opts.table, 'opts.table')
  return result(translate(shownTables(checkString(formula, 'formula'), opts.table, s.local.syn.list), s.invariant, s.local))
}

const SPACE = new Set([' ', '\t', '\n', '\r'])

/** The structured references of a formula as Excel shows them in a locale with the list separator `list` (`toLocal`); a formula that does not parse as it is. */
function shownTables(formula: string, table: string | undefined, list: string): string {
  if (!formula.includes('[')) return formula
  let toks: Tok[]
  try {
    parse(formula)
    toks = tokenize(formula)
  } catch (e) {
    if (e instanceof FormulaError) return formula
    throw e
  }
  let out = ''
  let pos = 0
  for (let k = 0; k < toks.length - 1; k++) {
    const tok = toks[k]
    if (tok.t !== 'table') continue
    let end = toks[k + 1].at
    while (end > tok.at && SPACE.has(formula[end - 1])) end--
    const own = table !== undefined && (tok.table === null || sheetKey(tok.table) === sheetKey(table))
    const name = own && unnamedInTable(tok.spec) ? '' : (tok.table ?? table ?? '')
    const text = !tok.spec.areas.length && !tok.spec.cols && name ? name : `${name}[${shownSpec(tok.spec, {}, list)}]`
    out += formula.slice(pos, tok.at) + text
    pos = end
  }
  return pos === 0 ? formula : out + formula.slice(pos)
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
