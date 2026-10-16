/**
 * The pieces of formula text that the tokenizer (`tokens.ts`), the locale lexer (`locale/lex.ts`) and the reader of a
 * file's formulas (`xlsx/formulas/from-file.ts`) scan alike: quoted and bracketed runs, whitespace, a call after a word,
 * a number, the characters of a word. What each does with a run left open stays with it. And text half typed, walked
 * outside its strings, quoted sheets' names and tables' brackets (`walk`), for `callAt`, `refAt` and a formula closed on Return.
 */

/** Whitespace between the tokens of a formula (an intersection where it stands between two references). */
export const isSpace = (ch: string | undefined) => ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r'

/**
 * The characters of `t` from `from` up to `to` outside strings, quoted sheets' names and tables' brackets, each given to
 * `f` with its place; `f` answering true stops the walk. What is open where it stopped: a string's or a quoted sheet's
 * quote, and tables' brackets.
 */
export function walk(t: string, from: number, to: number, f: (c: string, i: number) => boolean | void): { quote: string | null; table: number } {
  let quote: string | null = null
  let table = 0
  for (let i = from; i < to; i++) {
    const c = t[i]
    if (quote) {
      if (c === quote) quote = null
    } else if (c === '"' || c === "'") quote = c
    else if (c === '[') table++
    else if (c === ']') table = Math.max(0, table - 1)
    else if (!table && f(c, i)) break
  }
  return { quote, table }
}
/** A character of a function's name: a letter of any script, a digit, `.` or `_` (`СТАНДОТКЛОН.В`, `LOG10`). */
export const NAME_CHAR = /[\p{L}\p{N}._]/u

/** A bracket or brace left open in typed text: a function's, with its name, its argument's index so far and where it opened; a group's; an array constant's. */
export type OpenBracket = { kind: 'call'; name: string; index: number; open: number } | { kind: 'group' | 'array' }

/**
 * The brackets and braces typed formula text `t` leaves open at `to`, walked from its second character as `walk` walks
 * it, a call's arguments counted by the list separator `sep`; with what `walk` leaves open there, whether a `)` was met
 * and the names of the functions called.
 */
export function openBrackets(t: string, to: number, sep = ','): { open: OpenBracket[]; quote: string | null; table: number; shut: boolean; calls: string[] } {
  const open: OpenBracket[] = []
  const calls: string[] = []
  let shut = false
  const { quote, table } = walk(t, 1, to, (c, i) => {
    if (c === '{') open.push({ kind: 'array' })
    else if (c === '}') {
      if (open.at(-1)?.kind === 'array') open.pop()
    } else if (c === '(') {
      let from = i
      while (from > 1 && NAME_CHAR.test(t[from - 1])) from--
      if (from < i) calls.push(t.slice(from, i))
      open.push(from < i ? { kind: 'call', name: t.slice(from, i), index: 0, open: i } : { kind: 'group' })
    } else if (c === ')') {
      shut = true
      if (open.at(-1)?.kind !== 'array') open.pop()
    } else if (c === sep) {
      const top = open.at(-1)
      if (top?.kind === 'call') top.index++
    }
  })
  return { open, quote, table, shut, calls }
}

/** The index of the quote closing the run
 `src[i]` opens (`"…"` or `'…'`, the quote doubled inside), -1 when it is not closed. */
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
export const BOOK_SHEET = /^(?:[\p{L}\p{N}_.:]+|'(?:[^']|'')+')?!/u

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
 * A character of a word just before `[` or `_` in a file's formula, an identifier's (`?` too, as Excel loads a name
 * `a?_xlfn.b` and a table `T?`: the set `tables/name-question-held`): before `[` it makes the run a table's (`T?[Qty]`),
 * before `_` the word goes on and holds no prefix of a newer function (`a?_xlfn.b`).
 */
export const FILE_WORD_CHAR = new RegExp(IDENT_NEXT, 'u')
