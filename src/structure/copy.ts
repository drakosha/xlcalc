/**
 * Copy and paste, and fill: relative parts of the copied formulas move with the copy, absolute ones stay,
 * a reference pushed past the sheet's edge is `#REF!` (a range too: copies do not clamp). A target that
 * is a whole multiple of the source's size is tiled with copies, any other is pasted once at its top-left
 * cell, as Excel's Range.Copy does. Empty source cells clear the cells they land on.
 */
import type { BookInput, CellInput, CellRef, RestructureResult, TableInput } from '../contract.js'
import { spend } from '../exec/budget.js'
import { checkAddress, checkString, InputError } from '../input.js'
import { formatAddress, parseAddress } from '../model.js'
import { FormulaError, sheetKey, tokenize } from '../syntax.js'
import { tableAt, tablesOf } from '../tables/model.js'
import { rewriteTableRefs } from '../tables/rewrite.js'
import { isFull, LIMIT, span, type Area, type Axis, type Bound, type Rect } from './area.js'
import { carryNames, shiftRect, within } from './move.js'
import { arrayRefusal, copiedCells, relocated } from './arrays.js'
import { mergeRefusal, pasteLayout, type Paste } from './paste.js'
import { rewriteFormula, sheetNames, type Outcome, type Transform } from './rewrite.js'

/** Relative bounds moved by the copy's offset; a whole column keeps its rows, a whole row its columns. */
export function copyTransform(dr: number, dc: number): Transform {
  return {
    copy: true,
    ref({ sheet, area }): Outcome {
      const next: Area = { ...area }
      for (const [axis, d] of [['r', dr], ['c', dc]] as [Axis, number][]) {
        if (d === 0 || isFull(area[axis], axis)) continue
        const move = (b: Bound): Bound => (b.abs ? b : { i: b.i + d, abs: false })
        const s = span(move(area[axis].lo), move(area[axis].hi))
        if (s.lo.i < 0 || s.hi.i >= LIMIT[axis]) return 'lost'
        next[axis] = s
      }
      return { sheet, area: next }
    },
  }
}

/** The tiles of a paste: offsets of each copy of `src` in `target`. */
function tiles(src: Rect, target: Rect): [number, number][] {
  const [h, w] = [src.r2 - src.r1 + 1, src.c2 - src.c1 + 1]
  const [th, tw] = [target.r2 - target.r1 + 1, target.c2 - target.c1 + 1]
  if (th % h !== 0 || tw % w !== 0) return [[target.r1 - src.r1, target.c1 - src.c1]]
  const out: [number, number][] = []
  for (let r = target.r1; r <= target.r2; r += h) for (let c = target.c1; c <= target.c2; c += w) out.push([r - src.r1, c - src.c1])
  return out
}

/** A copied cell without its `meta` (rows of a source a host keeps by `meta` are not copied with the cell). */
const withoutMeta = (cell: CellInput): CellInput => {
  if (cell.meta === undefined) return cell
  const { meta: _, ...rest } = cell
  return rest
}

/**
 * Copy `src` of sheet `from` to `target` of sheet `to`; the copies keep `meta` only `withMeta`. The source is read
 * from `source`, the book as it was copied (the clipboard), when pastes follow one another.
 */
export function copyRange(book: BookInput, from: string, src: Rect, to: string, target: Rect, withMeta: boolean, source = book): RestructureResult {
  const offsets = tiles(src, target)
  const [dr, dc] = offsets[0]
  if (src.r2 + dr >= LIMIT.r || src.c2 + dc >= LIMIT.c) return { ok: false, code: 'off_sheet', message: `the copy does not fit on ${to} at ${formatAddress(target.r1, target.c1)}` }
  // the tiles cover the target, or one copy its top-left part
  const pasted = offsets.length > 1 ? target : shiftRect(src, dr, dc)
  const paste: Paste = { from, src, to, area: pasted, offsets, move: false }
  const refused = mergeRefusal(book, paste)
  if (refused) return { ok: false, code: 'merged', message: refused }
  const part = arrayRefusal(book, paste)
  if (part) return { ok: false, code: 'array', message: part }
  const [fk, tk] = [sheetKey(from), sheetKey(to)]
  const names = sheetNames(book)
  const sheet = source.sheets.find((s) => sheetKey(s.name) === fk)!
  // the cells of part of an array formula's area copy as array formulas of their own
  const parts = copiedCells(sheet, src, (f, r, c) => rewriteFormula(f, from, from, copyTransform(r, c), names))
  const inSrc = Object.keys(sheet.cells).filter((addr) => (spend(1), within(src, addr) && !parts.some(([a]) => a === addr)))
  const copied = [...inSrc.map((addr): [string, CellInput] => [addr, sheet.cells[addr]]), ...parts].map(([addr, cell]) => ({ addr, at: parseAddress(addr), cell: withMeta ? cell : withoutMeta(cell) }))
  const sheets = book.sheets.map((s) => {
    if (sheetKey(s.name) !== tk) return s
    const cells: Record<string, CellInput> = {}
    for (const addr of Object.keys(s.cells)) {
      const cell = s.cells[addr]
      spend(1)
      if (!within(pasted, addr)) cells[addr] = cell
    }
    for (const [r, c] of offsets) {
      const t = copyTransform(r, c)
      for (const { addr, at, cell } of copied) {
        // a paste over a large target stops at the session's deadline or abort (`exec/budget.ts`)
        spend(1)
        const formula = cell.formula === undefined ? undefined : rewriteFormula(cell.formula, to, to, t, names)
        const dest = formatAddress(at.row + r, at.col + c)
        cells[dest] = relocated(formula === cell.formula ? cell : { ...cell, formula }, addr, dest)
      }
    }
    return { ...s, cells }
  })
  const out: BookInput = { ...book, sheets }
  if (book.names) {
    const repoint: Transform = { ref: ({ sheet, area }) => (sheetKey(sheet) === fk ? { sheet: to, area } : undefined) }
    const formulas = copied.flatMap(({ cell }) => (cell.formula === undefined ? [] : [cell.formula]))
    const carried = fk === tk ? [] : carryNames(book.names, formulas, from, to, (ref) => rewriteFormula(ref, null, null, repoint, names))
    out.names = carried.length ? [...book.names, ...carried] : book.names
  }
  pasteLayout(source, out, paste)
  const mapCell = (ref: CellRef): CellRef | null => (sheetKey(ref.sheet) === tk && within(pasted, ref.cell) ? null : ref)
  return { ok: true, book: out, mapCell }
}

/** What `copyFormula` reads of a book: its sheets' names and their tables; a `BookInput` is one. */
export interface CopyContext {
  sheets: { name: string; tables?: TableInput[] }[]
}

/** The sheets a formula names, each as it spells it: without a book they are the sheets references point to. */
function namedSheets(text: string): Map<string, string> {
  const out = new Map<string, string>()
  try {
    for (const tok of tokenize(text)) if (tok.t === 'ref') for (const s of [tok.sheet, tok.last]) if (s) out.set(sheetKey(s), s)
  } catch (e) {
    if (!(e instanceof FormulaError)) throw e
  }
  return out
}

/** A cell of the book, checked: its sheet as the book spells it (any sheet without a book) and its row and column. */
function place(ref: unknown, path: string, book: CopyContext | undefined): { sheet: string; row: number; col: number } {
  if (typeof ref !== 'object' || ref === null) throw new InputError(path, 'expected a cell such as { sheet: "Sheet1", cell: "B2" }')
  const { sheet, cell } = ref as CellRef
  const name = checkString(sheet, `${path}.sheet`)
  const known = book?.sheets.find((s) => sheetKey(s.name) === sheetKey(name))
  if (book && !known) throw new InputError(`${path}.sheet`, `unknown sheet ${JSON.stringify(name)}`)
  return { sheet: known?.name ?? name, ...parseAddress(checkAddress(cell, `${path}.cell`)) }
}

/**
 * The text of a formula copied from cell `from` to cell `to`, as `copyRange` pastes it: relative parts move by the
 * offset, absolute ones stay, a reference pushed past the sheet's edge is `#REF!`, only the changed references are
 * rewritten. With the book (its sheets' names and tables), a formula copied out of a table names the table in its
 * references to it, as Excel pastes it; without it the formula's own sheet names are taken as written.
 */
export function copyFormula(text: string, from: CellRef, to: CellRef, book?: CopyContext): string {
  const formula = checkString(text, 'text')
  const [a, b] = [place(from, 'from', book), place(to, 'to', book)]
  const names = book ? sheetNames(book as BookInput) : namedSheets(formula)
  const out = rewriteFormula(formula, b.sheet, b.sheet, copyTransform(b.row - a.row, b.col - a.col), names)
  if (!book) return out
  const tables = tablesOf(book as BookInput)
  const index = (name: string) => book.sheets.findIndex((s) => sheetKey(s.name) === sheetKey(name))
  const t = tableAt(tables, index(a.sheet), a.row, a.col)
  // out of its table, as copyRange pastes it (`tables/moves.ts`): the references to the table without a name get its name
  return t && tableAt(tables, index(b.sheet), b.row, b.col) !== t ? rewriteTableRefs(out, tables, t, null, () => undefined, true) : out
}
