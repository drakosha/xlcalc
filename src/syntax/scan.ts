/**
 * The pieces of formula text that the tokenizer (`tokens.ts`), the locale lexer (`locale/lex.ts`) and the reader of a
 * file's formulas (`xlsx/formulas/from-file.ts`) scan alike: quoted and bracketed runs, whitespace, a call after a word,
 * a number, the characters of a word. What each does with a run left open stays with it.
 */

/** Whitespace between the tokens of a formula (an intersection where it stands between two references). */
export const isSpace = (ch: string | undefined) => ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r'

/** The index of the quote closing the run `src[i]` opens (`"…"` or `'…'`, the quote doubled inside), -1 when it is not closed. */
export function closingQuote(src: string, i: number): number {
  const q = src[i]
  for (let j = i + 1; j < src.length; j++) {
    if (src[j] !== q) continue
    if (src[j + 1] === q) j++
    else return j
  }
  return -1
}

/** The end of the bracketed run that opens at `i`, nested brackets and `'` escapes inside; the text's length when it is not closed. */
export function bracketEnd(src: string, i: number): number {
  let depth = 0
  for (let j = i; j < src.length; j++) {
    if (src[j] === "'") j++
    else if (src[j] === '[') depth++
    else if (src[j] === ']' && --depth === 0) return j + 1
  }
  return src.length
}

/** After a bracketed run, what makes it a reference to another workbook (`[1]Sheet1!A1`, `[Book.xlsx]S!A1`, `[1]!name`). */
export const BOOK_SHEET = /^(?:[\p{L}\p{N}_.]+|'(?:[^']|'')+')?!/u

/** Whether `(` follows at `i`, past spaces: the word before it is a function's name (`SUM (A1)`). */
export function callFollows(src: string, i: number): boolean {
  while (src[i] === ' ') i++
  return src[i] === '('
}

/** A number with the decimal separator `decimal` (a regular expression's source): `1.5`, `.5`, `5.`, `1.5E-07`. */
export const numberSource = (decimal: string) => `(?:[0-9]+(?:${decimal}[0-9]*)?|${decimal}[0-9]+)(?:[eE][+-]?[0-9]+)?`

// The characters of a word past its first differ by what the word is; each set below names where its difference comes from.

/**
 * An identifier, a function's or a defined name's: a letter, `_` or `\`, then letters, digits, `_`, `.`, `\` and `?`, as
 * Excel takes names (the sets `tables/names-more-held`, `names-ru-more-held`). A sheet's name without quotes takes no `\`
 * or `?` (`PLAIN_SHEET` in `sheet-pattern.ts`); the locale lexer reads the same identifiers, a name's check the same
 * characters (`IDENT_FIRST`, `IDENT_NEXT`).
 */
export const IDENT_FIRST = '[\\p{L}_\\\\]'
export const IDENT_NEXT = '[\\p{L}\\p{N}_.\\\\?]'
export const IDENT_RULE = `${IDENT_FIRST}${IDENT_NEXT}*`

/** A character that goes on with a word past a cell's shape, so the shape is no reference: an identifier's, or a call's `(` (`A1B`, `LOG10(`). */
export const IDENT_CHAR = /[\p{L}\p{N}_.(\\?]/u

/**
 * A character of a word just before `[` or `_` in a file's formula: before `[` a table's name (Excel's table names take no
 * `?`), before `_` the word goes on and holds no prefix of a newer function (`ab_xlfn.c`). It predates `?` in names: a
 * name `a?_xlfn.b` loses `_xlfn.` there, as before.
 */
export const FILE_WORD_CHAR = /[\p{L}\p{N}_.\\]/u
