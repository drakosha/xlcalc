/**
 * The references of formula text as it is typed (`scanReferences`), for a host that colours them: the engine's tokenizer
 * reads the text, and a token it does not read is skipped, so a formula half typed or broken in the middle gives the
 * references around the place. Text in a locale is read through its translation into stored text, whose positions map
 * back to the typed ones.
 */
import type { FormulaReference, ScanOptions } from './contract.js'
import { checkString, object, onlyKeys } from './input/basic.js'
import { sidesOf } from './locale/index.js'
import { optionalLocale } from './locale/load.js'
import { translate, type Translated } from './locale/translate.js'
import type { Locale } from './locale/types.js'
import { FormulaError, tokenize, type Tok } from './syntax.js'
import { bracketEnd, closingQuote, isSpace } from './syntax/scan.js'
import { AREA_WORDS } from './syntax/table.js'

/** The characters of a word the tokenizer reads as one token: a name's or a reference's (`$`, `:`, `!`, a spill's `#`), an error literal's. */
const WORD = /[\p{L}\p{N}_.\\?$:!#]/u

/**
 * Where the token at `at` that does not read ends: a quoted or bracketed run at its close, or at the end of the text when
 * it is open (a string being typed holds no references); a word at its last character, quoted and bracketed runs in it
 * too (`'Q1 data'!$A$0`); any other character alone.
 */
function skipped(text: string, at: number): number {
  let j = at
  while (j < text.length) {
    const ch = text[j]
    if (ch === '"' || ch === "'") {
      const q = closingQuote(text, j)
      j = q < 0 ? text.length : q + 1
    } else if (ch === '[') j = bracketEnd(text, j)
    else if (WORD.test(ch)) j++
    else break
  }
  return Math.max(j, Math.min(at + 1, text.length))
}

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

/**
 * The tokens of stored formula text in runs, each with the place it starts at: a token that does not read ends a run with
 * the tokens before it (an unreadable token's place can come after its start, past a sheet with nothing after it), and the
 * next run starts past it.
 */
function tokenRuns(text: string): { from: number; toks: Tok[] }[] {
  const out: { from: number; toks: Tok[] }[] = []
  for (let from = 0; from < text.length; ) {
    const rest = text.slice(from)
    try {
      out.push({ from, toks: tokenize(rest) })
      break
    } catch (e) {
      if (!(e instanceof FormulaError)) throw e
      out.push({ from, toks: tokensUpTo(rest.slice(0, e.at)) })
      from += skipped(rest, e.at)
    }
  }
  return out
}

/**
 * Typed text in a locale as stored text, with positions back in `text`: a piece that does not translate is blanked out
 * with spaces, which keep the positions and the array constant it stands in, and the text is translated again.
 */
function storedText(text: string, offset: number, locale: Locale): Translated {
  const { local, invariant } = sidesOf(locale)
  let typed = text
  for (;;) {
    const t = translate(typed.slice(offset), local, invariant, offset)
    if (t.ok) return t
    const end = skipped(typed, t.at)
    // nothing left to blank: the text before the place, shorter each time
    if (!/\S/.test(typed.slice(t.at, end))) typed = typed.slice(0, Math.min(t.at, typed.length - 1))
    else typed = typed.slice(0, t.at) + ' '.repeat(end - t.at) + typed.slice(end)
  }
}

const KIND = { cell: 'cells', range: 'cells', cols: 'columns', rows: 'rows' } as const

/**
 * The references to cells, entire columns and rows, and tables in formula text, in order, with their positions in it
 * (`at` to `end`), for coloring them as they are typed. `formula` is a stored formula, or text typed in `opts.locale`;
 * a leading `=` is skipped. A token that does not parse (a stray character, a half-typed error literal, a sheet name
 * with nothing after it) is skipped and parsing continues past it; an unclosed string, quoted sheet name or bracket
 * takes up the rest of the text. Strings, function names and defined names are not references.
 */
export function scanReferences(formula: string, opts?: ScanOptions): FormulaReference[] {
  const text = checkString(formula, 'formula')
  if (opts !== undefined) onlyKeys(object(opts, 'opts', "options such as { locale: 'ru-RU' }"), 'opts', ['locale'])
  const locale = optionalLocale(opts?.locale, 'opts.locale')
  const eq = text.startsWith('=') ? 1 : 0
  const stored = locale === undefined ? { formula: text.slice(eq), back: (at: number) => at + eq } : storedText(text, eq, locale)
  const out: FormulaReference[] = []
  for (const { from, toks } of tokenRuns(stored.formula))
    for (let k = 0; k < toks.length - 1; k++) {
      const tok = toks[k]
      if (tok.t !== 'ref' && tok.t !== 'table') continue
      // a reference ends where the next token begins, the whitespace before it left out
      let e = from + toks[k + 1].at
      while (e > from + tok.at && isSpace(stored.formula[e - 1])) e--
      const [at, end] = [stored.back(from + tok.at), stored.back(e)]
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
