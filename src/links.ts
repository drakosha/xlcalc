/**
 * Other workbooks' stored values in the engine (`BookInput.externalBooks`), as Excel computes a file it opens without
 * updating its links. Each linked sheet is a sheet of the runtime after the book's own, which no name of the book and
 * no edit reaches; another workbook's defined names are put in the place of their references before a formula is
 * compiled. A cell the stored values do not hold is blank, or `#REF!` on a sheet whose last update failed
 * (`refreshError`) or of which nothing is stored: there those cells are written as `#REF!` where formulas read them.
 */
import type { ExternalBook, ExternalValue } from './contract.js'
import { caseKey } from './collate.js'
import { put } from './evaluate.js'
import { MAX_COLS, parseAddress } from './model.js'
import type { SheetData } from './runtime.js'
import { FormulaError, linkedSheet as linkedSheetText, mapAst, parse, sheetKey, type Ast } from './syntax.js'
import { E, type Rng, type V } from './values.js'

/** The links of a book as the engine holds them. */
export interface LinkState {
  /** The index of the first linked sheet in the runtime: the book's own sheets come before. */
  first: number
  /** The index of each linked sheet by `book:sheet key`. */
  sheets: Map<string, number>
  /** Of each linked sheet by its index past `first`: the cells it stores (row · MAX_COLS + column) when a cell it does not store is `#REF!`, else null. */
  strict: (Set<number> | null)[]
  /** What each name of a linked book refers to, by `book:sheet:name key` (no sheet for a book's own): a reference, or null for `#REF!`. */
  names: Map<string, Ast | null>
  /** Each linked sheet by its index past `first` as a reference names it, unquoted (`[1]Sheet1`). */
  labels: string[]
}

const sheetAt = (book: number, sheet: string) => `${book}\u0000${sheetKey(sheet)}`
const nameAt = (book: number, sheet: number | undefined, name: string) => `${book}\u0000${sheet ?? ''}\u0000${caseKey(name)}`

/** A stored value as the runtime holds it. */
const runtimeValue = (v: ExternalValue): V => (typeof v === 'object' ? (E[v.error] ?? E.REF) : v)

/** What a linked book's name refers to: a reference into one of its sheets, or null (a constant or a lost reference, `#REF!` in Excel). */
function nameTarget(book: number, b: ExternalBook, ref: string): Ast | null {
  let ast: Ast
  try {
    ast = parse(ref.startsWith('=') ? ref.slice(1) : ref)
  } catch (e) {
    if (e instanceof FormulaError) return null
    throw e
  }
  if (ast.k !== 'ref' || ast.sheet === null || ast.ref.spill) return null
  const sheet = b.sheets.find((s) => sheetKey(s.name) === sheetKey(ast.sheet!))
  return sheet ? { k: 'xref', book, sheet: sheet.name, ref: ast.ref } : null
}

/** The linked sheets of a book as runtime sheets after its `first` own, and the links' state. */
export function linkSheets(books: readonly ExternalBook[] | undefined, first: number): { data: SheetData[]; links: LinkState } {
  const links: LinkState = { first, sheets: new Map(), strict: [], names: new Map(), labels: [] }
  const data: SheetData[] = []
  ;(books ?? []).forEach((b, i) => {
    const book = i + 1
    for (const s of b.sheets) {
      links.sheets.set(sheetAt(book, s.name), first + data.length)
      const sh: SheetData = { name: s.name, cols: [], rows: 0, fills: [] }
      const stored = s.cells === undefined || s.refreshError ? new Set<number>() : null
      for (const [addr, v] of Object.entries(s.cells ?? {})) {
        const { row, col } = parseAddress(addr)
        stored?.add(row * MAX_COLS + col)
        if (v !== null) put(sh, row, col, runtimeValue(v))
      }
      data.push(sh)
      links.strict.push(stored)
      links.labels.push(linkedSheetText(book, s.name))
    }
    for (const n of b.names ?? []) links.names.set(nameAt(book, n.sheet, n.name), nameTarget(book, b, n.ref))
  })
  return { data, links }
}

/** A runtime sheet's name as a reference writes it, unquoted: the book's own, or a linked one's (`[1]Sheet1`). */
export const runtimeSheet = (sheetNames: readonly string[], links: LinkState, s: number): string => (s < links.first ? sheetNames[s] : links.labels[s - links.first])

/** The runtime index of a linked sheet, undefined when the book has no such link or sheet. */
export const linkedSheet = (links: LinkState, book: number, sheet: string): number | undefined => links.sheets.get(sheetAt(book, sheet))

/**
 * The formula with the names of linked books replaced by what they refer to (`#REF!` for a name the book does not have):
 * `[1]Sheet1!Local` is the sheet's own name, else the book's, as a sheet sees a book's names.
 */
export function expandLinkNames(ast: Ast, links: LinkState, books: readonly ExternalBook[] | undefined): Ast {
  if (!links.names.size && !books?.length) return ast
  return mapAst(ast, (n) => {
    if (n.k !== 'xname' || typeof n.book !== 'number') return undefined
    const j = n.sheet === undefined ? undefined : books?.[n.book - 1]?.sheets.findIndex((s) => sheetKey(s.name) === sheetKey(n.sheet!))
    const target = (j !== undefined && j >= 0 ? links.names.get(nameAt(n.book, j, n.name)) : undefined) ?? links.names.get(nameAt(n.book, undefined, n.name))
    return target ?? { k: 'err', code: 'REF' }
  })
}

/** At most this many cells of a sheet are written as `#REF!` for the ranges formulas read, so a whole column stays cheap. */
const FILL_LIMIT = 2_000_000

/**
 * The cells of strict linked sheets that the ranges `refs` read and the stored values do not hold, to be `#REF!`. A
 * range is read down to one row past the sheet's stored rows and across to one column past its stored columns (a
 * whole column of a sheet stores no more), its first cell at least.
 */
export function strictCells(links: LinkState, sheets: readonly SheetData[], refs: Iterable<Rng>): [s: number, r: number, c: number][] {
  const out: [number, number, number][] = []
  const seen = new Map<number, Set<number>>()
  for (const g of refs) {
    const stored = links.strict[g.s - links.first]
    if (g.s < links.first || !stored) continue
    const sh = sheets[g.s]
    const r2 = Math.max(g.r1, Math.min(g.r2, sh.rows))
    const c2 = Math.max(g.c1, Math.min(g.c2, sh.cols.length))
    let done = seen.get(g.s)
    if (!done) seen.set(g.s, (done = new Set()))
    for (let r = g.r1; r <= r2; r++)
      for (let c = g.c1; c <= c2; c++) {
        const k = r * MAX_COLS + c
        if (stored.has(k) || done.has(k) || done.size >= FILL_LIMIT) continue
        done.add(k)
        out.push([g.s, r, c])
      }
  }
  return out
}
