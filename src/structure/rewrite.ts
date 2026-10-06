/**
 * Rewriting the references of formula text in place: only the text of a changed reference is replaced,
 * the rest of the formula keeps its spelling. A formula that does not parse is left as written.
 */
import type { BookInput, CellInput, NameInput } from '../contract.js'
import { FormulaError, parse, quoteSheet, sheetKey, tokenize, type Tok } from '../syntax.js'
import { decode, encode, sameArea, type Area } from './area.js'

/** A reference as a transform sees it: the sheet it points to (as the book spells it) and its rectangle. */
export interface Site {
  sheet: string
  area: Area
}

/** Unchanged, a new place, lost on its sheet (`#REF!`, `'S'!#REF!`), or on a deleted sheet (`#REF!A1`). */
export type Outcome = undefined | { sheet: string; area: Area } | 'lost' | 'gone'

export interface Transform {
  ref(site: Site): Outcome
  /** The new name of a sheet in `'S'!#REF!`: null when the sheet is deleted, undefined to keep it. */
  sheet?(name: string): string | null | undefined
}

/** Sheet names of a book by key (sheet names compare case-insensitively). */
export const sheetNames = (book: BookInput) => new Map(book.sheets.map((s) => [sheetKey(s.name), s.name]))

const same = (a: string | null, b: string | null) => a !== null && b !== null && sheetKey(a) === sheetKey(b)

const SPACE = new Set([' ', '\t', '\n', '\r'])

/**
 * The formula with its references rewritten by `t`. `host` is the formula's sheet before the operation,
 * `moved` after it (a moved formula's references without a sheet get the sheet they point to); null for
 * the reference of a name, which always has its sheet.
 */
export function rewriteFormula(text: string, host: string | null, moved: string | null, t: Transform, sheets: Map<string, string>): string {
  try {
    parse(text)
  } catch (e) {
    if (e instanceof FormulaError) return text
    throw e
  }
  const toks = tokenize(text)
  let out = ''
  let pos = 0
  for (let k = 0; k < toks.length - 1; k++) {
    const tok = toks[k]
    let end = toks[k + 1].at
    while (end > tok.at && SPACE.has(text[end - 1])) end--
    const now = tok.t === 'ref' ? refText(tok, text.slice(tok.at, end), host, moved, t, sheets) : tok.t === 'err' && tok.sheet !== undefined ? errText(tok.sheet, t, sheets) : undefined
    if (now === undefined) continue
    out += text.slice(pos, tok.at) + now
    pos = end
  }
  return pos === 0 ? text : out + text.slice(pos)
}

/** The new text of a reference token, undefined when it stays as written. */
function refText(tok: Extract<Tok, { t: 'ref' }>, written: string, host: string | null, moved: string | null, t: Transform, sheets: Map<string, string>): string | undefined {
  const target = tok.sheet === null ? host : (sheets.get(sheetKey(tok.sheet)) ?? null)
  if (target === null) return undefined
  const area = decode(tok.ref)
  const o = t.ref({ sheet: target, area })
  // without its sheet a reference points to the formula's sheet: a formula moved elsewhere spells it out
  const prefix = (sheet: string) => (tok.sheet !== null && sheet === target ? written.slice(0, written.length - tok.ref.text.length) : tok.sheet !== null || !same(sheet, moved ?? host) ? `${quoteSheet(sheet)}!` : '')
  if (o === undefined) return tok.sheet === null && !same(target, moved ?? host) ? prefix(target) + tok.ref.text : undefined
  if (o === 'gone') return `#REF!${tok.ref.text}`
  if (o === 'lost') return `${prefix(target)}#REF!`
  const next = prefix(o.sheet) + (sameArea(o.area, area) ? tok.ref.text : encode(o.area))
  return next === written ? undefined : next
}

function errText(name: string, t: Transform, sheets: Map<string, string>): string | undefined {
  const next = t.sheet?.(sheets.get(sheetKey(name)) ?? name)
  if (next === undefined) return undefined
  return next === null ? '#REF!#REF!' : `${quoteSheet(next)}!#REF!`
}

/** A cell with its formula rewritten; the same object when nothing changed. */
export function rewriteCell(cell: CellInput, host: string, moved: string, t: Transform, sheets: Map<string, string>): CellInput {
  if (cell.formula === undefined) return cell
  const formula = rewriteFormula(cell.formula, host, moved, t, sheets)
  return formula === cell.formula ? cell : { ...cell, formula }
}

/** The references of names rewritten; the same array when nothing changed. */
export function rewriteNames(names: NameInput[] | undefined, t: Transform, sheets: Map<string, string>): NameInput[] | undefined {
  if (!names) return names
  let changed = false
  const out = names.map((n) => {
    const ref = rewriteFormula(n.ref, null, null, t, sheets)
    if (ref === n.ref) return n
    changed = true
    return { ...n, ref }
  })
  return changed ? out : names
}

/**
 * Every formula of the book (and the names) rewritten in place, for operations that move no cells;
 * `rename` gives a sheet's new name.
 */
export function rewriteBook(book: BookInput, t: Transform, sheets: Map<string, string>, rename = (name: string) => name): BookInput {
  const out: BookInput = {
    ...book,
    sheets: book.sheets.map((s) => {
      const name = rename(s.name)
      let cells: Record<string, CellInput> | undefined
      for (const [addr, cell] of Object.entries(s.cells)) {
        const next = rewriteCell(cell, s.name, name, t, sheets)
        if (next !== cell) (cells ??= { ...s.cells })[addr] = next
      }
      return cells || name !== s.name ? { ...s, name, cells: cells ?? s.cells } : s
    }),
  }
  const names = rewriteNames(book.names, t, sheets)
  return names ? { ...out, names } : out
}
