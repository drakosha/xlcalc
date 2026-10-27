/**
 * Copy and paste, and fill: relative parts of the copied formulas move with the copy, absolute ones stay,
 * a reference pushed past the sheet's edge is `#REF!` (a range too: copies do not clamp). A target that
 * is a whole multiple of the source's size is tiled with copies, any other is pasted once at its top-left
 * cell, as Excel's Range.Copy does. Empty source cells clear the cells they land on.
 */
import type { BookInput, CellInput, CellRef, TableInput } from '../contract.js'
import { spend } from '../exec/budget.js'
import { InputError } from '../input-error.js'
import { checkString } from '../input/basic.js'
import { checkSheetOf } from '../input.js'
import { checkAddress } from '../input/range.js'
import { say } from '../locale/messages.js'
import { formatAddress, parseAddress } from '../model.js'
import { FormulaError, sheetKey, tokenize } from '../syntax.js'
import { tableAt, tablesOf } from '../tables/model.js'
import { rewriteTableRefs } from '../tables/rewrite.js'
import { LIMIT, shiftRect, type Rect } from './area.js'
import { carryNames } from './move.js'
import { copiedCells, relocated } from './arrays.js'
import { placer, type Applied, type Geometry } from './geometry.js'
import { cellsIn, gridOf } from './grid.js'
import { pasteLayout, pasteRefusal, type Paste } from './paste.js'
import { formulaTokens, rewriteFormula, sheetIndex, sheetNames, sheetOf, type Transform } from './rewrite.js'
import { copyTransform } from './transform.js'
import { linksIn, markedLinks } from '../syntax/links.js'
import { copyConditional } from '../cond/structure.js'
import { pasteValidations } from '../validation/structure.js'
import { pastedLooks } from '../looks/move.js'
import { layerOf, withLayer } from '../looks/sheet.js'

/** The copies of `src` a paste lays in `target`: as many as fit when the target is a whole multiple of the source, else one. */
export function tileCount(src: Rect, target: Rect): number {
  const [h, w] = [src.r2 - src.r1 + 1, src.c2 - src.c1 + 1]
  const [th, tw] = [target.r2 - target.r1 + 1, target.c2 - target.c1 + 1]
  return th % h !== 0 || tw % w !== 0 ? 1 : (th / h) * (tw / w)
}

/**
 * The tiles of a paste, offsets of each copy of `src` in `target`, and the area they cover: the target when it is a
 * whole multiple of the source, else one copy at its top-left cell.
 */
export function tiles(src: Rect, target: Rect): { offsets: [number, number][]; area: Rect } {
  if (tileCount(src, target) === 1) return { offsets: [[target.r1 - src.r1, target.c1 - src.c1]], area: shiftRect(src, target.r1 - src.r1, target.c1 - src.c1) }
  const [h, w] = [src.r2 - src.r1 + 1, src.c2 - src.c1 + 1]
  const offsets: [number, number][] = []
  for (let r = target.r1; r <= target.r2; r += h) for (let c = target.c1; c <= target.c2; c += w) offsets.push([r - src.r1, c - src.c1])
  return { offsets, area: target }
}

/**
 * A copied cell without its `held`, as a formula pasted is one entered (`CellInput.held`), and with `meta` only `withMeta`
 * (rows of a source a host keeps by `meta` are not copied with the cell).
 */
const pastedCell = (cell: CellInput, withMeta: boolean): CellInput => {
  if (cell.held === undefined && (withMeta || cell.meta === undefined)) return cell
  const { meta, held: _, ...rest } = cell
  return withMeta && meta !== undefined ? { ...rest, meta } : rest
}

/**
 * Copy `src` of sheet `from` to `target` of sheet `to`; the copies keep `meta` only `withMeta`. The source is read
 * from `source`, the book as it was copied (the clipboard), when pastes follow one another. A copy of which a paste
 * takes the looks alone (`looksOnly`) may cover part of an array formula.
 */
export function copyRange(book: BookInput, from: string, src: Rect, to: string, target: Rect, withMeta: boolean, source = book, looksOnly = false): Applied {
  const { offsets, area: pasted } = tiles(src, target)
  const [dr, dc] = offsets[0]
  if (src.r2 + dr >= LIMIT.r || src.c2 + dc >= LIMIT.c) return { ok: false, code: 'off_sheet', message: say(book.locale, { key: 'copyOffSheet', sheet: to, cell: formatAddress(target.r1, target.c1) }) }
  const paste: Paste = { from, src, to, area: pasted, offsets, move: false }
  const refusal = pasteRefusal(book, paste, looksOnly)
  if (refusal) return refusal
  const [fk, tk] = [sheetKey(from), sheetKey(to)]
  const ti = sheetIndex(book, to)
  const geometry: Geometry = [{ kind: 'clear', s: ti, area: pasted }]
  const names = sheetNames(book)
  const sheet = sheetOf(source, from)
  // the cells of part of an array formula's area copy as array formulas of their own
  const parts = copiedCells(sheet, src, (f, r, c) => rewriteFormula(f, from, from, copyTransform(r, c), names))
  const own = new Set(parts.map(([addr]) => addr))
  const g = gridOf(sheet)
  const copied = [
    ...cellsIn(g, src).flatMap((i) => (own.has(g.addr[i]) ? [] : [{ addr: g.addr[i], at: { row: g.row[i], col: g.col[i] }, cell: g.cell[i] }])),
    ...parts.map(([addr, cell]) => ({ addr, at: parseAddress(addr), cell })),
  ].map((x) => {
    const cell = pastedCell(x.cell, withMeta)
    return cell === x.cell ? x : { ...x, cell }
  })
  const sheets = book.sheets.map((s, si) => {
    if (si !== ti) return s
    // the target's cells but those the geometry takes away, the pasted ones over them
    const cells: Record<string, CellInput> = {}
    const [t, place, p] = [gridOf(s), placer(geometry), { s: ti, r: 0, c: 0 }]
    for (let i = 0; i < t.addr.length; i++) {
      spend(1)
      p.r = t.row[i]
      p.c = t.col[i]
      if (place(p)) cells[t.addr[i]] = t.cell[i]
    }
    // each copied formula read once for all its copies
    const read = copied.map(({ cell }) => (cell.formula === undefined ? null : formulaTokens(cell.formula, false)))
    for (const [r, c] of offsets) {
      const t = copyTransform(r, c)
      for (let k = 0; k < copied.length; k++) {
        const { addr, at, cell } = copied[k]
        // a paste over a large target stops at the session's deadline or abort (`exec/budget.ts`)
        spend(1)
        const formula = cell.formula === undefined ? undefined : rewriteFormula(cell.formula, to, to, t, names, read[k])
        const dest = formatAddress(at.row + r, at.col + c)
        cells[dest] = relocated(formula === cell.formula ? cell : { ...cell, formula }, addr, dest)
      }
    }
    // the looks of the source's cells without an entry, laid by rectangles over the tiles
    const looks = pastedLooks(layerOf(sheet), src, layerOf(s), pasted)
    return looks ? withLayer(s, cells, looks) : { ...s, cells }
  })
  const out: BookInput = { ...book, sheets }
  // formulas pasted with links: Excel marks their books as it cannot read them (`markedLinks`)
  const linked = markedLinks(book.externalBooks, copied.flatMap(({ cell }) => (cell.formula === undefined ? [] : linksIn(cell.formula))))
  if (linked !== book.externalBooks) out.externalBooks = linked
  if (book.names) {
    const repoint: Transform = { ref: ({ sheet, area }) => (sheetKey(sheet) === fk ? { sheet: to, area } : undefined) }
    const formulas = copied.flatMap(({ cell }) => (cell.formula === undefined ? [] : [cell.formula]))
    const carried = fk === tk ? [] : carryNames(book.names, formulas, from, to, (ref) => rewriteFormula(ref, null, null, repoint, names))
    out.names = carried.length ? [...book.names, ...carried] : book.names
  }
  pasteValidations(source, out, paste, null, names)
  pasteLayout(source, out, paste)
  copyConditional(source, out, paste, names)
  return { ok: true, book: out, geometry }
}

/** What `copyFormula` reads from a workbook: its sheet names and their tables; a `BookInput` qualifies. */
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
  const known = book ? checkSheetOf(book, sheet, `${path}.sheet`) : undefined
  return { sheet: known?.name ?? checkString(sheet, `${path}.sheet`), ...parseAddress(checkAddress(cell, `${path}.cell`)) }
}

/** A cell by its sheet's name, its row and its column, zero-based. */
interface Place {
  sheet: string
  row: number
  col: number
}

/** The formula `formula` copied from `a` to `b` on `b`'s sheet, its references read with the sheets `names`. */
const copied = (formula: string, a: Place, b: Place, names: Map<string, string>) => rewriteFormula(formula, b.sheet, b.sheet, copyTransform(b.row - a.row, b.col - a.col), names)

/**
 * `copyFormula` over a book for many copies, its sheets' names and its tables (`tables`, those of its session) read once:
 * the text of a formula copied from `a` to `b`, cells of the book's sheets as it spells them.
 */
export function formulaCopier(book: CopyContext, tables = tablesOf(book as BookInput)): (formula: string, a: Place, b: Place) => string {
  const names = sheetNames(book as BookInput)
  return (formula, a, b) => {
    const out = copied(formula, a, b, names)
    const t = tableAt(tables, sheetIndex(book, a.sheet), a.row, a.col)
    // out of its table, as copyRange pastes it (`tables/moves.ts`): the references to the table without a name get its name
    return t && tableAt(tables, sheetIndex(book, b.sheet), b.row, b.col) !== t ? rewriteTableRefs(out, tables, t, null, () => undefined, true) : out
  }
}

/**
 * The text of a formula copied from cell `from` to cell `to`, as `copyRange` pastes it: relative parts shift by the
 * offset, absolute ones stay, a reference pushed beyond the sheet's edge becomes `#REF!`, and only the changed
 * references are rewritten. With the workbook (its sheet names and tables), a formula copied out of a table names the
 * table in its references to it, as Excel pastes it; without it, the sheet names in the formula are taken as written.
 */
export function copyFormula(text: string, from: CellRef, to: CellRef, book?: CopyContext): string {
  const formula = checkString(text, 'text')
  const [a, b] = [place(from, 'from', book), place(to, 'to', book)]
  return book ? formulaCopier(book)(formula, a, b) : copied(formula, a, b, namedSheets(formula))
}
