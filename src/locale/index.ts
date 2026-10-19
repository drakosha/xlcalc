/**
 * Locales of formula text: typing and showing formulas the way Excel does in another language, while books
 * keep the invariant text. `toLocal` shows a stored formula, `toInvariant` reads typed text back, and
 * `validateLocalFormula` checks typed text with positions in it; their messages are in the locale's language. Each
 * takes the tag of a loaded locale or a locale object, which needs no loading.
 */
import type { ExternalBook, ValidateContext, ValidateResult } from '../contract.js'
import { checkExternalBooks, checkValidateContext } from '../input.js'
import { checkString, object, onlyOptions } from '../input/basic.js'
import { FormulaError, parse, sheetKey, validate, type Tok } from '../syntax.js'
import { eachToken, numberLinks, showLinks } from '../syntax/links.js'
import { shownSpec, unnamedInTable } from '../syntax/table.js'
import { EN_US } from '../locales/en-US.js'
import { checkLocaleArg } from './load.js'
import { say } from './messages.js'
import { sides, translate, type Fail, type Sides, type Translated } from './translate.js'
import type { Locale, Translation } from './types.js'

export type { Locale, Translation } from './types.js'

/** en-US formula text, the invariant syntax; the package's other locales are separate modules (`xlcalc/locales/ru-RU`). */
export const enUS: Locale = EN_US

const cache = new WeakMap<object, Sides>()

/** A checked locale's translation tables (`checkLocaleArg`), built once per locale object (a locale must not change after its first use). */
export function sidesOf(locale: Locale): Sides {
  let s = cache.get(locale)
  if (!s) cache.set(locale, (s = sides(locale)))
  return s
}

/** A locale given by a loaded locale's tag or as an object, and its translation tables. */
function localeSides(where: string | Locale): [Locale, Sides] {
  const locale = checkLocaleArg(where, 'locale')
  return [locale, sidesOf(locale)]
}

/** A translation's result, its message in the locale's language. */
const result = (r: Translated | Fail, locale: Locale): Translation => (r.ok ? { ok: true, formula: r.formula } : refusal(r, locale))
const refusal = (r: Fail, locale: Locale) => ({ ok: false as const, code: 'syntax' as const, message: say(locale, r.msg), at: r.at })

/**
 * A stored formula as displayed in a locale (a loaded locale's tag, or a locale object): `SUM(A1:A3,1.5)` becomes
 * `СУММ(A1:A3;1,5)` in ru-RU, and a sheet name is quoted only where Excel quotes it (`'Data'!B2` becomes `Data!B2`, in
 * en-US too; `quoteSheetIfNeeded` with the locale's booleans). Structured references are displayed as Excel displays
 * them, the formula's own row as `@` (`Sales[[#This Row],[Qty]]` becomes `Sales[@Qty]`) and an entire table by its name
 * alone (`Sales[]` becomes `Sales`); `opts.table`, the table containing the formula's cell, drops that table's name
 * before its own columns (`[@Qty]`, `SUM([Qty])`), as Excel displays a formula inside a table.
 */
export function toLocal(formula: string, where: string | Locale, opts?: TranslateOptions): Translation {
  const sides = localeSides(where)
  const { table, externalBooks } = checkTranslateOptions(opts)
  return localText(checkString(formula, 'formula'), sides, table, externalBooks)
}

/**
 * `toLocal` for the package's own calls, which give it a checked book's table and links, without checking them: a
 * window and an explanation show each formula of a book whose links can hold thousands of sheets.
 */
export const showFormula = (formula: string, where: string | Locale, table?: string, externalBooks?: ExternalBook[]): Translation => localText(formula, localeSides(where), table, externalBooks)

function localText(formula: string, [locale, s]: [Locale, Sides], table: string | undefined, externalBooks: ExternalBook[] | undefined): Translation {
  const shown = showLinks(shownTables(formula, table, s.local.syn.list), externalBooks ?? [])
  return result(translate(shown, s.invariant, s.local), locale)
}

/**
 * `toLocal`'s and `toInvariant`'s options: `table`, the table of the formula's cell; `externalBooks`, the book's links to
 * other workbooks, by which a reference to one is shown as Excel shows it (`'[1]Sheet1'!A1` is `[Book2.xlsx]Sheet1!A1`,
 * `'../data/[Book3.xlsx]Sheet1'!A1` with the path the link holds) and read back by its file name.
 */
export interface TranslateOptions {
  table?: string
  externalBooks?: ExternalBook[]
}

function checkTranslateOptions(v: unknown): TranslateOptions {
  if (v === undefined) return {}
  const o = object(v, 'opts', "options such as { table: 'Sales' }")
  onlyOptions(o, 'opts', ['table', 'externalBooks'])
  if (o.table !== undefined) checkString(o.table, 'opts.table')
  if (o.externalBooks !== undefined) checkExternalBooks(o.externalBooks, 'opts.externalBooks', false)
  return o as TranslateOptions
}

/** The structured references of a formula as Excel shows them in a locale with the list separator `list` (`toLocal`); a formula that does not parse as it is. */
function shownTables(formula: string, table: string | undefined, list: string): string {
  if (!formula.includes('[')) return formula
  try {
    parse(formula)
  } catch (e) {
    if (e instanceof FormulaError) return formula
    throw e
  }
  return eachTable(formula, (tok) => {
    const own = table !== undefined && (tok.table === null || sheetKey(tok.table) === sheetKey(table))
    const name = own && unnamedInTable(tok.spec) ? '' : (tok.table ?? table ?? '')
    return !tok.spec.areas.length && !tok.spec.cols && name ? name : `${name}[${shownSpec(tok.spec, list)}]`
  })
}

/** `formula` with the text of each structured reference as `f` writes it; a formula that does not read as it is. */
const eachTable = (formula: string, f: (tok: Extract<Tok, { t: 'table' }>, text: string) => string): string =>
  formula.includes('[') ? eachToken(formula, (tok, text) => (tok.t === 'table' ? f(tok, text) : undefined)) : formula

/**
 * Text typed in a locale (a loaded locale's tag, or a locale object) as a stored formula: `=СУММ(A1:A3;1,5)` becomes
 * `SUM(A1:A3,1.5)`. A leading `=` is removed, since stored formulas have none; English function names, TRUE/FALSE and
 * error texts are also accepted, and a sheet name with or without quotes is kept as typed. Error positions refer to
 * `text`, and error messages are in the locale's language. `opts.table`, the table containing the formula's cell, adds
 * that table's name before a structured reference without one (`[@Qty]` becomes `Sales[@Qty]`), the reverse of
 * `toLocal` dropping it, so the formula also works outside the table.
 */
export function toInvariant(text: string, where: string | Locale, opts?: TranslateOptions): Translation {
  const [locale, s] = localeSides(where)
  const { table, externalBooks } = checkTranslateOptions(opts)
  const eq = checkString(text, 'text').startsWith('=') ? 1 : 0
  const t = translate(text.slice(eq), s.local, s.invariant, eq)
  if (!t.ok) return refusal(t, locale)
  const formula = table === undefined ? t.formula : eachTable(t.formula, (tok, x) => (tok.table === null ? table + x : x))
  if (!externalBooks) return { ok: true, formula }
  const n = numberLinks(formula, externalBooks)
  return n.ok ? n : { ok: false, code: 'external_ref', message: say(locale, n.message), at: t.back(n.at) }
}

/**
 * `validateFormula` for text typed in a locale (a loaded locale's tag, or a locale object): the same result, with the
 * canonical stored text, error positions referring to `text` (a leading `=` is allowed) and messages in the locale's
 * language.
 */
export function validateLocalFormula(text: string, where: string | Locale, ctx?: ValidateContext): ValidateResult {
  const sides = localeSides(where)
  checkString(text, 'text')
  return validatedText(text, sides, checkValidateContext(ctx))
}

/** `validateLocalFormula` for the package's own calls, which give it a checked context: typed text checks a book's links for each formula. */
export const validateTyped = (text: string, where: string | Locale, ctx?: ValidateContext): ValidateResult => validatedText(text, localeSides(where), ctx)

function validatedText(text: string, [locale, s]: [Locale, Sides], c: ValidateContext | undefined): ValidateResult {
  const eq = text.startsWith('=') ? 1 : 0
  const t = translate(text.slice(eq), s.local, s.invariant, eq)
  if (!t.ok) return refusal(t, locale)
  const r = validate(t.formula, { ...c, say: (m) => say(locale, m) })
  if (!r.ok) return { ...r, at: t.back(r.at) }
  const { ast: _ast, ...rest } = r
  return rest
}
