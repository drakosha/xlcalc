/**
 * Locales of formula text: typing and showing formulas the way Excel does in another language, while books
 * keep the invariant text. `toLocal` shows a stored formula, `toInvariant` reads typed text back, and
 * `validateLocalFormula` checks typed text with positions in it; their messages are in the locale's language. Each
 * takes the tag of a loaded locale or a locale object, which needs no loading.
 */
import type { ValidateContext, ValidateResult } from '../contract.js'
import { checkString, checkValidateContext } from '../input.js'
import { FormulaError, parse, sheetKey, tokenize, validate, type Tok } from '../syntax.js'
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

/** The translation tables of a locale given by its tag or as an object, built once per locale object. */
function sidesOf(where: string | Locale): [Locale, Sides] {
  const locale = checkLocaleArg(where, 'locale')
  let s = cache.get(locale)
  if (!s) cache.set(locale, (s = sides(locale)))
  return [locale, s]
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
export function toLocal(formula: string, where: string | Locale, opts: { table?: string } = {}): Translation {
  const [locale, s] = sidesOf(where)
  if (opts.table !== undefined) checkString(opts.table, 'opts.table')
  return result(translate(shownTables(checkString(formula, 'formula'), opts.table, s.local.syn.list), s.invariant, s.local), locale)
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
 * Text typed in a locale (a loaded locale's tag, or a locale object) as a stored formula: `=СУММ(A1:A3;1,5)` is
 * `SUM(A1:A3,1.5)`. A leading `=` is dropped, as stored formulas have none; English function names, TRUE/FALSE and
 * error texts are read too. Positions of errors are in `text`, their messages in the locale's language.
 */
export function toInvariant(text: string, where: string | Locale): Translation {
  const [locale, s] = sidesOf(where)
  const eq = checkString(text, 'text').startsWith('=') ? 1 : 0
  return result(translate(text.slice(eq), s.local, s.invariant, eq), locale)
}

/**
 * `validateFormula` for text typed in a locale (a loaded locale's tag, or a locale object): the same result, with the
 * canonical stored text, positions of errors in `text` (a leading `=` is allowed) and messages in the locale's language.
 */
export function validateLocalFormula(text: string, where: string | Locale, ctx?: ValidateContext): ValidateResult {
  const [locale, s] = sidesOf(where)
  const eq = checkString(text, 'text').startsWith('=') ? 1 : 0
  const c = checkValidateContext(ctx)
  const t = translate(text.slice(eq), s.local, s.invariant, eq)
  if (!t.ok) return refusal(t, locale)
  const r = validate(t.formula, { ...c, say: (m) => say(locale, m) })
  if (!r.ok) return { ...r, at: t.back(r.at) }
  const { ast: _ast, ...rest } = r
  return rest
}
