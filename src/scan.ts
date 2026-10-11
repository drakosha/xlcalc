/**
 * The references of formula text as it is typed (`scanReferences`), for a host that colours them: the engine's tokenizer
 * reads the text as far as it goes, so a formula half typed gives the references before the place it stops reading.
 * Text in a locale is read through its translation into stored text, whose positions map back to the typed ones.
 */
import type { FormulaReference, ScanOptions } from './contract.js'
import { checkString, object, onlyKeys } from './input.js'
import { sidesOf } from './locale/index.js'
import { checkLocaleArg } from './locale/load.js'
import { translate } from './locale/translate.js'
import type { Locale } from './locale/types.js'
import { FormulaError, tokenize, type Tok } from './syntax.js'
import { AREA_WORDS } from './syntax/table.js'

/** The tokens of stored formula text up to the first place it does not read (all of them when it reads). */
function tokensUpTo(text: string): Tok[] {
  let end = text.length
  for (;;) {
    try {
      return tokenize(text.slice(0, end))
    } catch (e) {
      if (!(e instanceof FormulaError)) throw e
      end = Math.min(e.at, end - 1)
    }
  }
}

/** Typed text in a locale as stored text up to the first place it does not translate, with positions back in `text`. */
function storedUpTo(text: string, offset: number, locale: Locale): { formula: string; back(at: number): number } {
  const { local, invariant } = sidesOf(locale)
  let end = text.length - offset
  for (;;) {
    const t = translate(text.slice(offset, offset + end), local, invariant, offset)
    if (t.ok) return t
    end = Math.min(t.at - offset, end - 1)
  }
}

const KIND = { cell: 'cells', range: 'cells', cols: 'columns', rows: 'rows' } as const
const SPACE = /\s/

/**
 * The references to cells, whole columns and rows, and tables in formula text, in order, with their place in it (`at` to
 * `end`), for colouring them as they are typed. `formula` is a stored formula, or typed text in `opts.locale`; a leading
 * `=` is skipped. Text that does not read yet (a formula half typed, an open string) gives the references before the
 * place it stops reading; strings, function names and defined names are no references.
 */
export function scanReferences(formula: string, opts?: ScanOptions): FormulaReference[] {
  const text = checkString(formula, 'formula')
  if (opts !== undefined) onlyKeys(object(opts, 'opts', "options such as { locale: 'ru-RU' }"), 'opts', ['locale'])
  const locale = opts?.locale === undefined ? undefined : checkLocaleArg(opts.locale, 'opts.locale')
  const eq = text.startsWith('=') ? 1 : 0
  const stored = locale === undefined ? { formula: text.slice(eq), back: (at: number) => at + eq } : storedUpTo(text, eq, locale)
  const toks = tokensUpTo(stored.formula)
  const out: FormulaReference[] = []
  for (let k = 0; k < toks.length - 1; k++) {
    const tok = toks[k]
    if (tok.t !== 'ref' && tok.t !== 'table') continue
    // a reference ends where the next token begins, the whitespace before it left out
    let e = toks[k + 1].at
    while (e > tok.at && SPACE.test(stored.formula[e - 1])) e--
    const [at, end] = [stored.back(tok.at), stored.back(e)]
    const place = { at, end, text: text.slice(at, end) }
    if (tok.t === 'table') {
      const { table, spec } = tok
      out.push({ kind: 'table', ...place, ...(table === null ? {} : { table }), ...(spec.cols ? { columns: spec.cols } : {}), ...(spec.areas.length ? { items: spec.areas.map((a) => AREA_WORDS[a]) } : {}) })
      continue
    }
    const { ref } = tok
    const sheets = tok.sheet === null ? {} : { sheet: tok.sheet, ...(tok.last === undefined ? {} : { sheetLast: tok.last }) }
    out.push({ kind: KIND[ref.kind], ...place, ...sheets, top: ref.r1 + 1, left: ref.c1 + 1, bottom: ref.r2 + 1, right: ref.c2 + 1, ...(ref.spill ? { spill: true } : {}) })
  }
  return out
}
