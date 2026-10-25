/**
 * References to other workbooks bound to a book's links (`BookInput.externalBooks`): a link's number as a file writes
 * it (`[1]Sheet1!A1`), or the file name a person reads and types (`[Book2.xlsx]Sheet1!A1`, `'C:\data\[Book2.xlsx]Sheet1'!A1`,
 * `Book2.xlsx!Rate`) found among the links by the last part of its path. The canonical text numbers each.
 */
import type { CheckCode, ExternalBook } from '../contract.js'
import { caseKey } from '../collate.js'
import type { Message } from '../locale/messages.js'
import { PLAIN_NAME } from '../sheet-pattern.js'
import { FormulaError, type Ast } from './ast.js'
import { mapAst, quoteSheet, sheetKey } from './print.js'
import { isSpace } from './scan.js'
import { readsAsCell, tokenize, type Tok } from './tokens.js'

/** A link's target as a formula shows it: the folder (`C:\data\`, `../data/`, empty for a file name) and the file. */
export function shownTarget(target: string): { dir: string; file: string } {
  const t = target.replace(/^file:\/\/\/?/i, '')
  const k = Math.max(t.lastIndexOf('/'), t.lastIndexOf('\\')) + 1
  return { dir: t.slice(0, k), file: t.slice(k) }
}

/** The number of the link a path or file name names (from 1), the whole path first, then the file alone; 0 for none. */
export function linkOf(books: readonly ExternalBook[], path: string): number {
  const whole = caseKey(path.replace(/^file:\/\/\/?/i, ''))
  const exact = books.findIndex((b) => caseKey(b.target.replace(/^file:\/\/\/?/i, '')) === whole)
  if (exact >= 0) return exact + 1
  const file = caseKey(shownTarget(path).file)
  return books.findIndex((b) => caseKey(shownTarget(b.target).file) === file) + 1
}

/** Whether a sheet's name in a reference is no sheet of the book but a path or a file name a link has (`Book2.xlsx!Rate`). */
const namesLink = (books: readonly ExternalBook[], sheet: string, sheets: Set<string> | null) => !sheets?.has(sheetKey(sheet)) && /\.\w+$/.test(sheet) && linkOf(books, sheet) > 0

export type LinkFail = { code: CheckCode; message: Message; needle: string }

/**
 * The formula's references to other workbooks numbered by `books` and their sheets spelled as the links hold them; a
 * name seen from a sheet that is a link's file (`Book2.xlsx!Rate`, `'C:\data\Book2.xlsx'!Rate`) becomes that book's name.
 * A reference to a book the links do not hold, or to a sheet its link does not have, fails (`external_ref`).
 */
export function bindLinks(ast: Ast, books: readonly ExternalBook[] | undefined, sheets: readonly string[] | undefined): { ok: true; ast: Ast } | ({ ok: false } & LinkFail) {
  const own = sheets ? new Set(sheets.map(sheetKey)) : null
  let fail: LinkFail | null = null
  const number = (book: number | string): number => {
    const k = typeof book === 'number' ? (books && book <= books.length ? book : 0) : books ? linkOf(books, book) : 0
    if (!k && !fail) fail = { code: 'external_ref', message: books ? { key: 'unknownLink', book: typeof book === 'number' ? `[${book}]` : book } : { key: 'externalReference' }, needle: typeof book === 'number' ? `[${book}]` : shownTarget(book).file }
    return k
  }
  const sheetOf = (k: number, sheet: string): string | undefined => {
    const s = books![k - 1].sheets.find((x) => sheetKey(x.name) === sheetKey(sheet))?.name
    if (s === undefined && !fail) fail = { code: 'external_ref', message: { key: 'unknownLinkSheet', book: shownTarget(books![k - 1].target).file, sheet }, needle: sheet }
    return s
  }
  const out = mapAst(ast, (n) => {
    if (fail) return undefined
    if (n.k === 'name' && n.sheet !== undefined && books?.length && namesLink(books, n.sheet, own)) return { k: 'xname', book: linkOf(books, n.sheet), name: n.name }
    if (n.k === 'err' && n.book !== undefined) {
      const k = number(n.book)
      return k && k !== n.book ? { ...n, book: k } : undefined
    }
    if (n.k !== 'xref' && n.k !== 'xname') return undefined
    const k = number(n.book)
    if (!k) return undefined
    if (n.sheet === undefined) return n.book === k ? undefined : { ...n, book: k }
    const sheet = sheetOf(k, n.sheet)
    if (sheet === undefined) return undefined
    return n.book === k && n.sheet === sheet ? undefined : { ...n, book: k, sheet }
  })
  return fail ? { ok: false, ...(fail as LinkFail) } : { ok: true, ast: out }
}

/** A name of a sheet or a file in a reference as Excel writes it: in quotes unless it is a plain name (`Book2.xlsx`, `Sheet1`). */
export const plain = (s: string) => PLAIN_NAME.test(s) && !readsAsCell(s)

/** Another workbook's sheet before `!` as Excel shows it, by the link's target: `[Book2.xlsx]Sheet1`, `'C:\data\[Book2.xlsx]My Sheet'`. */
export function shownLinkSheet(target: string, sheet: string): string {
  const { dir, file } = shownTarget(target)
  const text = `${dir}[${file}]${sheet}`
  return dir === '' && plain(file) && plain(sheet) ? text : quoteSheet(text)
}

/** Another workbook's own name as Excel shows it: `Book2.xlsx!Rate`, `'C:\data\Book2.xlsx'!Rate`. */
export function shownLinkName(target: string, name: string): string {
  const { dir, file } = shownTarget(target)
  return `${dir === '' && plain(file) ? file : quoteSheet(dir + file)}!${name}`
}

/**
 * Each token of `formula` that `f` gives a text (undefined keeps it) replaced by that text; a formula that does not read
 * as it is. `given`: its tokens read already.
 */
export function eachToken(formula: string, f: (tok: Tok, text: string) => string | undefined, given?: Tok[]): string {
  let toks: Tok[]
  try {
    toks = given ?? tokenize(formula)
  } catch (e) {
    if (e instanceof FormulaError) return formula
    throw e
  }
  let out = ''
  let pos = 0
  for (let k = 0; k < toks.length - 1; k++) {
    const tok = toks[k]
    let end = toks[k + 1].at
    while (end > tok.at && isSpace(formula[end - 1])) end--
    const text = f(tok, formula.slice(tok.at, end))
    if (text === undefined) continue
    out += formula.slice(pos, tok.at) + text
    pos = end
  }
  return pos === 0 ? formula : out + formula.slice(pos)
}

/** A stored formula's references to other workbooks as Excel shows them, by the links' targets (`[1]` is `[Book2.xlsx]`); a number the links lack as it is. */
export function showLinks(formula: string, books: readonly ExternalBook[]): string {
  if (!books.length || !formula.includes('[')) return formula
  return eachToken(formula, (tok) => {
    if ((tok.t !== 'xref' && tok.t !== 'xname' && tok.t !== 'err') || typeof tok.book !== 'number' || !books[tok.book - 1]) return undefined
    const target = books[tok.book - 1].target
    if (tok.t === 'err') return `${shownLinkSheet(target, tok.sheet!)}!#REF!`
    if (tok.t === 'xref') return `${shownLinkSheet(target, tok.sheet)}!${tok.ref.text}`
    return tok.sheet === undefined ? shownLinkName(target, tok.v) : `${shownLinkSheet(target, tok.sheet)}!${tok.v}`
  })
}

/**
 * Typed text's references to other workbooks by the number of the link whose file they name (`[Book2.xlsx]Sheet1!A1`
 * is `'[1]Sheet1'!A1`, `Book2.xlsx!Rate` is `[1]!Rate`), or the first one no link holds.
 */
export function numberLinks(formula: string, books: readonly ExternalBook[], sheets?: readonly string[]): { ok: true; formula: string } | { ok: false; at: number; message: Message } {
  let miss: { at: number; message: Message } | null = null
  const own = sheets ? new Set(sheets.map(sheetKey)) : null
  const out = eachToken(formula, (tok) => {
    if (miss) return undefined
    if (tok.t === 'ident' && tok.sheet !== undefined && namesLink(books, tok.sheet, own)) return `[${linkOf(books, tok.sheet)}]!${tok.v}`
    if ((tok.t !== 'xref' && tok.t !== 'xname' && tok.t !== 'err') || tok.book === undefined || typeof tok.book === 'number') return undefined
    const k = linkOf(books, tok.book)
    if (!k) {
      miss = { at: tok.at, message: { key: 'unknownLink', book: tok.book } }
      return undefined
    }
    if (tok.t === 'err') return `${quoteSheet(`[${k}]${tok.sheet}`)}!#REF!`
    if (tok.t === 'xref') return `${quoteSheet(`[${k}]${tok.sheet}`)}!${tok.ref.text}`
    return tok.sheet === undefined ? `[${k}]!${tok.v}` : `${quoteSheet(`[${k}]${tok.sheet}`)}!${tok.v}`
  })
  return miss ? { ok: false, ...(miss as { at: number; message: Message }) } : { ok: true, formula: out }
}

/** The numbers of the links a formula's text refers to; none for text that does not read. */
export function linksIn(formula: string): number[] {
  if (!formula.includes('[')) return []
  const out = new Set<number>()
  try {
    for (const tok of tokenize(formula)) if ((tok.t === 'xref' || tok.t === 'xname') && typeof tok.book === 'number') out.add(tok.book)
  } catch (e) {
    if (!(e instanceof FormulaError)) throw e
  }
  return [...out]
}

/**
 * The links with every sheet of the books numbered `ks` marked as Excel marks them when a formula referring to a closed
 * book is entered (`refreshError`: it cannot read the book, and a cell the stored values do not hold is `#REF!` since);
 * the same array when nothing changes.
 */
export function markedLinks(books: ExternalBook[] | undefined, ks: Iterable<number>): ExternalBook[] | undefined {
  if (!books) return books
  let out: ExternalBook[] | undefined
  for (const k of ks) {
    const b = (out ?? books)[k - 1]
    if (!b || b.sheets.every((s) => s.refreshError)) continue
    out ??= [...books]
    out[k - 1] = { ...b, sheets: b.sheets.map((s) => (s.refreshError ? s : { ...s, refreshError: true })) }
  }
  return out ?? books
}
