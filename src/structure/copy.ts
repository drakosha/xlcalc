/**
 * Copy and paste, and fill: relative parts of the copied formulas move with the copy, absolute ones stay,
 * a reference pushed past the sheet's edge is `#REF!` (a range too: copies do not clamp). A target that
 * is a whole multiple of the source's size is tiled with copies, any other is pasted once at its top-left
 * cell, as Excel's Range.Copy does. Empty source cells clear the cells they land on.
 */
import type { BookInput, CellInput, CellRef } from '../contract.js'
import { formatAddress, parseAddress } from '../model.js'
import { sheetKey } from '../syntax.js'
import { isFull, LIMIT, span, type Area, type Axis, type Bound, type Rect } from './area.js'
import { carryNames, shiftRect, within } from './move.js'
import { rewriteFormula, sheetNames, type Outcome, type Transform } from './rewrite.js'
import type { RestructureResult } from './types.js'

/** Relative bounds moved by the copy's offset; a whole column keeps its rows, a whole row its columns. */
function copyTransform(dr: number, dc: number): Transform {
  return {
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

/** Copy `src` of sheet `from` to `target` of sheet `to`. */
export function copyRange(book: BookInput, from: string, src: Rect, to: string, target: Rect): RestructureResult {
  const offsets = tiles(src, target)
  const [dr, dc] = offsets[0]
  if (src.r2 + dr >= LIMIT.r || src.c2 + dc >= LIMIT.c) return { ok: false, code: 'off_sheet', message: `the copy does not fit on ${to} at ${formatAddress(target.r1, target.c1)}` }
  const [fk, tk] = [sheetKey(from), sheetKey(to)]
  const names = sheetNames(book)
  const source = Object.entries(book.sheets.find((s) => sheetKey(s.name) === fk)!.cells).filter(([addr]) => within(src, addr))
  const pasted = offsets.map(([r, c]) => shiftRect(src, r, c))
  const sheets = book.sheets.map((s) => {
    if (sheetKey(s.name) !== tk) return s
    const cells: Record<string, CellInput> = {}
    for (const [addr, cell] of Object.entries(s.cells)) if (!pasted.some((x) => within(x, addr))) cells[addr] = cell
    for (const [r, c] of offsets) {
      const t = copyTransform(r, c)
      for (const [addr, cell] of source) {
        const p = parseAddress(addr)
        const formula = cell.formula === undefined ? undefined : rewriteFormula(cell.formula, to, to, t, names)
        cells[formatAddress(p.row + r, p.col + c)] = formula === cell.formula ? cell : { ...cell, formula }
      }
    }
    return { name: s.name, cells }
  })
  const out: BookInput = { sheets }
  if (book.names) {
    const repoint: Transform = { ref: ({ sheet, area }) => (sheetKey(sheet) === fk ? { sheet: to, area } : undefined) }
    const formulas = source.flatMap(([, c]) => (c.formula === undefined ? [] : [c.formula]))
    const carried = fk === tk ? [] : carryNames(book.names, formulas, from, to, (ref) => rewriteFormula(ref, null, null, repoint, names))
    out.names = carried.length ? [...book.names, ...carried] : book.names
  }
  const mapCell = (ref: CellRef): CellRef | null => (sheetKey(ref.sheet) === tk && pasted.some((x) => within(x, ref.cell)) ? null : ref)
  return { ok: true, book: out, mapCell }
}
