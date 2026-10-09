/**
 * Rewriting the references of formula text in place: only the text of a changed reference is replaced,
 * the rest of the formula keeps its spelling. A formula that does not parse is left as written.
 */
import type { BookInput, CellInput, NameInput, TableColumn, TableInput } from '../contract.js'
import { spend } from '../exec/budget.js'
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
  /**
   * The new ends of a reference across sheets (`'Jan:Mar'!A1`): undefined to keep them, null when none is left; ends
   * on one sheet are written as a reference to it. Rows and columns of its sheets do not move it.
   */
  ends?(first: string, last: string): [string, string] | null | undefined
  /** A copy: relative references move by its offset wherever they point, across sheets too. */
  copy?: true
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
  // a session's operation stops here at its deadline or abort (`exec/budget.ts`); without a budget one comparison
  spend(1)
  try {
    // a name's formula may be a union
    parse(text, { union: host === null })
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
    const written = text.slice(tok.at, end)
    const now =
      tok.t === 'ref' ? (tok.last === undefined ? refText(tok, written, host, moved, t, sheets) : cubeText(tok, tok.last, written, t, sheets)) : tok.t === 'err' && tok.sheet !== undefined ? errText(tok.sheet, tok.ref ?? '', t, sheets) : undefined
    if (now === undefined) continue
    out += text.slice(pos, tok.at) + now
    pos = end
  }
  return pos === 0 ? text : out + text.slice(pos)
}

/**
 * The new text of a reference token, undefined when it stays as written. A spill reference (`E2#`) moves as
 * its cell and keeps its `#`, lost too (`#REF!#`, as Excel writes it).
 */
function refText(tok: Extract<Tok, { t: 'ref' }>, written: string, host: string | null, moved: string | null, t: Transform, sheets: Map<string, string>): string | undefined {
  const target = tok.sheet === null ? host : (sheets.get(sheetKey(tok.sheet)) ?? null)
  if (target === null) return undefined
  const area = decode(tok.ref)
  const o = t.ref({ sheet: target, area })
  const spill = tok.ref.spill ? '#' : ''
  // without its sheet a reference points to the formula's sheet: a formula moved elsewhere spells it out
  const prefix = (sheet: string) => (tok.sheet !== null && sheet === target ? written.slice(0, written.length - tok.ref.text.length - spill.length) : tok.sheet !== null || !same(sheet, moved ?? host) ? `${quoteSheet(sheet)}!` : '')
  if (o === undefined) return tok.sheet === null && !same(target, moved ?? host) ? prefix(target) + tok.ref.text + spill : undefined
  if (o === 'gone') return `#REF!${tok.ref.text}${spill}`
  if (o === 'lost') return `${prefix(target)}#REF!${spill}`
  const next = prefix(o.sheet) + (sameArea(o.area, area) ? tok.ref.text : encode(o.area)) + spill
  return next === written ? undefined : next
}

/** The new text of a reference across sheets, undefined when it stays as written. */
function cubeText(tok: Extract<Tok, { t: 'ref' }>, last: string, written: string, t: Transform, sheets: Map<string, string>): string | undefined {
  const spelled = (s: string) => sheets.get(sheetKey(s)) ?? s
  const [first, end] = [spelled(tok.sheet!), spelled(last)]
  const ends = t.ends?.(first, end)
  let area = decode(tok.ref)
  if (t.copy) {
    const o = t.ref({ sheet: first, area })
    if (o === 'lost' || o === 'gone') return `${quoteSheet(`${first}:${end}`)}!#REF!`
    if (o) area = o.area
  }
  if (ends === null) return `#REF!${encode(area)}`
  const [a, b] = ends ?? [first, end]
  const next = `${quoteSheet(sheetKey(a) === sheetKey(b) ? a : `${a}:${b}`)}!${encode(area)}`
  return next === written || (ends === undefined && sameArea(area, decode(tok.ref))) ? undefined : next
}

/** `'S'!#REF!` (`tail`: `#` of a lost spill reference) on a renamed or deleted sheet. */
function errText(name: string, tail: string, t: Transform, sheets: Map<string, string>): string | undefined {
  const next = t.sheet?.(sheets.get(sheetKey(name)) ?? name)
  if (next === undefined) return undefined
  return (next === null ? '#REF!#REF!' : `${quoteSheet(next)}!#REF!`) + tail
}

/** A cell with its formula rewritten; the same object when nothing changed. */
export function rewriteCell(cell: CellInput, host: string, moved: string, t: Transform, sheets: Map<string, string>): CellInput {
  // every cell of the loops over a book counts, a value too
  spend(1)
  if (cell.formula === undefined) return cell
  const formula = rewriteFormula(cell.formula, host, moved, t, sheets)
  return formula === cell.formula ? cell : { ...cell, formula }
}

/** A transform for names' formulas under inserted or deleted rows and columns: a reference without `$`, an offset from the cell using the name, stays. */
export const offsetsStay = (t: Transform): Transform => ({ ...t, ref: (site) => (site.area.r.lo.abs || site.area.r.hi.abs || site.area.c.lo.abs || site.area.c.hi.abs ? t.ref(site) : undefined) })

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
 * The formulas of a sheet's calculated columns rewritten as the formulas of their first data rows are: `host` is the
 * sheet before the operation, `moved` gives a table's sheet after it. The same array when nothing changed.
 */
export function rewriteTemplates(tables: TableInput[] | undefined, host: string, moved: (table: TableInput) => string, t: Transform, sheets: Map<string, string>): TableInput[] | undefined {
  if (!tables) return tables
  let out: TableInput[] | undefined
  tables.forEach((table, i) => {
    let columns: TableColumn[] | undefined
    table.columns.forEach((c, j) => {
      if (c.formula === undefined) return
      const formula = rewriteFormula(c.formula, host, moved(table), t, sheets)
      if (formula !== c.formula) (columns ??= [...table.columns])[j] = { ...c, formula }
    })
    if (columns) (out ??= [...tables])[i] = { ...table, columns }
  })
  return out ?? tables
}

/**
 * Every formula of the book (the names and the calculated columns too) rewritten in place, for operations that move
 * no cells; `rename` gives a sheet's new name.
 */
export function rewriteBook(book: BookInput, t: Transform, sheets: Map<string, string>, rename = (name: string) => name): BookInput {
  const out: BookInput = {
    ...book,
    sheets: book.sheets.map((s) => {
      const name = rename(s.name)
      let cells: Record<string, CellInput> | undefined
      for (const addr of Object.keys(s.cells)) {
        const cell = s.cells[addr]
        const next = rewriteCell(cell, s.name, name, t, sheets)
        if (next !== cell) (cells ??= { ...s.cells })[addr] = next
      }
      const tables = rewriteTemplates(s.tables, s.name, () => name, t, sheets)
      return cells || name !== s.name || tables !== s.tables ? { ...s, name, cells: cells ?? s.cells, ...(tables ? { tables } : {}) } : s
    }),
  }
  const names = rewriteNames(book.names, t, sheets)
  return names ? { ...out, names } : out
}
