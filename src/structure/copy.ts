/**
 * Copy and paste, and fill: relative parts of the copied formulas move with the copy, absolute ones stay,
 * a reference pushed past the sheet's edge is `#REF!` (a range too: copies do not clamp). A target that
 * is a whole multiple of the source's size is tiled with copies, any other is pasted once at its top-left
 * cell, as Excel's Range.Copy does. Empty source cells clear the cells they land on.
 */
import type { BookInput, CellInput, CellRef, RestructureResult } from '../contract.js'
import { formatAddress, parseAddress } from '../model.js'
import { sheetKey } from '../syntax.js'
import { isFull, LIMIT, span, type Area, type Axis, type Bound, type Rect } from './area.js'
import { carryNames, shiftRect, within } from './move.js'
import { arrayRefusal, copiedCells, relocated } from './arrays.js'
import { mergeRefusal, pasteLayout, type Paste } from './paste.js'
import { rewriteFormula, sheetNames, type Outcome, type Transform } from './rewrite.js'

/** Relative bounds moved by the copy's offset; a whole column keeps its rows, a whole row its columns. */
function copyTransform(dr: number, dc: number): Transform {
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
  const copied = [...Object.entries(sheet.cells).filter(([addr]) => within(src, addr) && !parts.some(([a]) => a === addr)), ...parts].map(([addr, cell]) => ({ addr, at: parseAddress(addr), cell: withMeta ? cell : withoutMeta(cell) }))
  const sheets = book.sheets.map((s) => {
    if (sheetKey(s.name) !== tk) return s
    const cells: Record<string, CellInput> = {}
    for (const [addr, cell] of Object.entries(s.cells)) if (!within(pasted, addr)) cells[addr] = cell
    for (const [r, c] of offsets) {
      const t = copyTransform(r, c)
      for (const { addr, at, cell } of copied) {
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
