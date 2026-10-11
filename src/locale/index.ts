/**
 * Locales of formula text: typing and showing formulas the way Excel does in another language, while books
 * keep the invariant text. `toLocal` shows a stored formula, `toInvariant` reads typed text back, and
 * `validateLocalFormula` checks typed text with positions in it; their messages are in the locale's language. Each
 * takes the tag of a loaded locale or a locale object, which needs no loading.
 */
import type { ValidateContext, ValidateResult } from '../contract.js'
import { checkString, checkValidateContext, object, onlyOptions } from '../input.js'
import { FormulaError, parse, sheetKey, tokenize, validate, type Tok } from '../syntax.js'
import { isSpace } from '../syntax/scan.js'
import { shownSpec, unnamedInTable } from '../syntax/table.js'
import { EN_US } from '../locales/en-US.js'
import { checkLocaleArg } from './load.js'
import { say } from './messages.js'
import { sides, translate, type Fail, type Sides, type Translated } from './translate.js'
import type { Locale, Translation } from './types.js'

export type { Locale, Translation } from './types.js'

/** en-US's formula text, the invariant syntax; the package's other locales are modules (`xlcalc/locales/ru-RU`). */
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
 * A stored formula as a locale shows it (a loaded locale's tag, or a locale object): `SUM(A1:A3,1.5)` is
 * `СУММ(A1:A3;1,5)` in ru-RU. Structured references are shown as Excel shows them, the formula's row as `@`
 * (`Sales[[#This Row],[Qty]]` is `Sales[@Qty]`) and a table alone by its name (`Sales[]` is `Sales`); `opts.table`, the
 * table of the formula's cell, drops its name before a column of its own (`[@Qty]`, `SUM([Qty])`), as Excel shows a
 * formula inside a table.
 */
export function toLocal(formula: string, where: string | Locale, opts?: { table?: string }): Translation {
  const [locale, s] = localeSides(where)
  const table = opts === undefined ? undefined : checkTranslateOptions(opts)
  return result(translate(shownTables(checkString(formula, 'formula'), table, s.local.syn.list), s.invariant, s.local), locale)
}

/** `toLocal`'s and `toInvariant`'s options: `table`, a table's name, alone; the table it gives. */
function checkTranslateOptions(v: unknown): string | undefined {
  const o = object(v, 'opts', "options such as { table: 'Sales' }")
  onlyOptions(o, 'opts', ['table'])
  return o.table === undefined ? undefined : checkString(o.table, 'opts.table')
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
function eachTable(formula: string, f: (tok: Extract<Tok, { t: 'table' }>, text: string) => string): string {
  if (!formula.includes('[')) return formula
  let toks: Tok[]
  try {
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
    while (end > tok.at && isSpace(formula[end - 1])) end--
    out += formula.slice(pos, tok.at) + f(tok, formula.slice(tok.at, end))
    pos = end
  }
  return pos === 0 ? formula : out + formula.slice(pos)
}

/**
 * Text typed in a locale (a loaded locale's tag, or a locale object) as a stored formula: `=СУММ(A1:A3;1,5)` is
 * `SUM(A1:A3,1.5)`. A leading `=` is dropped, as stored formulas have none; English function names, TRUE/FALSE and
 * error texts are read too. Positions of errors are in `text`, their messages in the locale's language. `opts.table`, the
 * table of the formula's cell, names it before a structured reference without a table name (`[@Qty]` is `Sales[@Qty]`),
 * as `toLocal` drops it, so the formula reads outside the table too.
 */
export function toInvariant(text: string, where: string | Locale, opts?: { table?: string }): Translation {
  const [locale, s] = localeSides(where)
  const table = opts === undefined ? undefined : checkTranslateOptions(opts)
  const eq = checkString(text, 'text').startsWith('=') ? 1 : 0
  const r = result(translate(text.slice(eq), s.local, s.invariant, eq), locale)
  return r.ok && table !== undefined ? { ok: true, formula: eachTable(r.formula, (tok, t) => (tok.table === null ? table + t : t)) } : r
}

/**
 * `validateFormula` for text typed in a locale (a loaded locale's tag, or a locale object): the same result, with the
 * canonical stored text, positions of errors in `text` (a leading `=` is allowed) and messages in the locale's language.
 */
export function validateLocalFormula(text: string, where: string | Locale, ctx?: ValidateContext): ValidateResult {
  const [locale, s] = localeSides(where)
  const eq = checkString(text, 'text').startsWith('=') ? 1 : 0
  const c = checkValidateContext(ctx)
  const t = translate(text.slice(eq), s.local, s.invariant, eq)
  if (!t.ok) return refusal(t, locale)
  const r = validate(t.formula, { ...c, say: (m) => say(locale, m) })
  if (!r.ok) return { ...r, at: t.back(r.at) }
  const { ast: _ast, ...rest } = r
  return rest
}
