/**
 * Cut and paste. A reference inside the moved range follows it; one inside the overwritten cells is
 * `#REF!`. A range with one end among the moved or the overwritten cells is adjusted by rules fitted to
 * Excel on every range of a 12-row window under 20 moves (`test/excel/struct/moves.json`, see
 * `docs/structure.md`). Across sheets a range only sheds the moved cells at its end.
 */
import type { BookInput, CellInput, NameInput } from '../contract.js'
import { caseKey } from '../collate.js'
import { say } from '../locale/messages.js'
import { formatAddress, MAX_COLS, MAX_ROWS } from '../model.js'
import { inputNameFor } from '../names.js'
import { parse, sheetKey, walk } from '../syntax.js'
import { contains, hasCell, hi, isFull, lo, meets, parseRange, rectOf, shiftRect, withSpan, type Area, type Axis, type Rect } from './area.js'
import { relocated } from './arrays.js'
import { placer, type Applied, type Geometry, type Place } from './geometry.js'
import { gridOf } from './grid.js'
import { pasteLayout, pasteRefusal, type Paste } from './paste.js'
import { rewriteCell, rewriteCells, rewriteNames, rewriteTemplates, sheetIndex, sheetNames, type Outcome, type Transform } from './rewrite.js'
import { moveConditional } from '../cond/structure.js'
import { pasteValidations } from '../validation/structure.js'
import { movedLooks } from '../looks/move.js'
import { layerOf, withLayer } from '../looks/sheet.js'

type Pair = [number, number]

/**
 * One axis of a range [i, j] when the cells [s1, s2] move by d along it, the moved cells covering the
 * range's other axis; 'ref' when the range lies in the overwritten cells.
 */
export function moveSpan(i: number, j: number, s1: number, s2: number, d: number): Pair | 'ref' {
  const d1 = s1 + d
  const d2 = s2 + d
  const inS = (x: number) => x >= s1 && x <= s2
  if (i >= s1 && j <= s2) return [i + d, j + d]
  if (i >= d1 && j <= d2 && (j < s1 || i > s2)) return 'ref'
  if (!inS(i) && !inS(j)) return overwrite(i, j, d1, d2)
  const top = inS(i)
  // an end moving away from the other end follows the cells
  if (top ? d < 0 : d > 0) return top ? [i + d, j] : [i, j + d]
  // an end moving inward past the other end: unchanged
  if (d2 < i || d1 > j) return [i, j]
  const shed: Pair = top ? [s2 + 1, j] : [i, s1 - 1]
  // the range held all the moved cells: it keeps the rest and takes in where they land
  if (s1 >= i && s2 <= j) return [Math.min(shed[0], d1), Math.max(shed[1], d2)]
  // the moved cells straddle the end: it follows them unless that leaves it empty or inside the
  // overwritten cells of a move that does not overlap itself; then it sheds them
  const follow: Pair = top ? [i + d, j] : [i, j + d]
  const overlap = d1 <= s2 && s1 <= d2
  const base = follow[0] < follow[1] && (overlap || follow[0] < d1 || follow[1] > d2) ? follow : shed
  return d > 0 ? [base[0], Math.max(base[1], d2)] : [Math.min(base[0], d1), base[1]]
}

/** A range [i, j] with an end among the overwritten cells [d1, d2] sheds them, unless it holds them all. */
export function overwrite(i: number, j: number, d1: number, d2: number): Pair {
  if (d1 >= i && d2 <= j) return [i, j]
  return [i >= d1 && i <= d2 ? d2 + 1 : i, j >= d1 && j <= d2 ? d1 - 1 : j]
}

const other = (axis: Axis): Axis => (axis === 'r' ? 'c' : 'r')

/** The axis along which `x` holds one end of `range` across the range's whole width, if any. */
function endAxis(range: Rect, x: Rect): Axis | null {
  for (const axis of ['r', 'c'] as const) {
    const o = other(axis)
    if (lo(x, o) > lo(range, o) || hi(x, o) < hi(range, o)) continue
    const a = lo(range, axis) >= lo(x, axis) && lo(range, axis) <= hi(x, axis)
    const b = hi(range, axis) >= lo(x, axis) && hi(range, axis) <= hi(x, axis)
    if (a !== b) return axis
  }
  return null
}

/** The area with `axis` set to `s`: undefined when that changes nothing; a range over the whole axis keeps it. */
function adjusted(sheet: string, area: Area, axis: Axis, s: Pair | 'ref'): Outcome {
  if (isFull(area[axis], axis)) return undefined
  if (s === 'ref') return 'lost'
  if (s[0] === area[axis].lo.i && s[1] === area[axis].hi.i) return undefined
  return { sheet, area: withSpan(area, axis, s[0], s[1]) }
}

/** A reference meeting the overwritten cells `dst`: `#REF!` inside them unless it holds moved cells, else shed at its end. */
function overwritten(sheet: string, area: Area, x: Rect, dst: Rect, holdsMoved: boolean): Outcome {
  if (contains(dst, x)) return holdsMoved ? undefined : 'lost'
  const axis = endAxis(x, dst)
  return axis && area.kind !== 'cell' ? adjusted(sheet, area, axis, overwrite(lo(x, axis), hi(x, axis), lo(dst, axis), hi(dst, axis))) : undefined
}

function moveTransform(from: string, src: Rect, to: string, dr: number, dc: number): Transform {
  const [fk, tk] = [sheetKey(from), sheetKey(to)]
  const dst = shiftRect(src, dr, dc)
  return {
    ref({ sheet, area }): Outcome {
      const x = rectOf(area)
      const k = sheetKey(sheet)
      if (k === fk && contains(src, x)) return { sheet: to, area: withSpan(withSpan(area, 'r', x.r1 + dr, x.r2 + dr), 'c', x.c1 + dc, x.c2 + dc) }
      const onSrc = k === fk && meets(x, src)
      const onDst = k === tk && meets(x, dst)
      if (!onSrc && !onDst) return undefined
      if (fk !== tk) {
        if (onDst) return overwritten(sheet, area, x, dst, false)
        // the moved cells leave the sheet: a range sheds them at its end
        const axis = endAxis(x, src)
        if (!axis) return undefined
        return adjusted(sheet, area, axis, lo(x, axis) >= lo(src, axis) ? [hi(src, axis) + 1, hi(x, axis)] : [lo(x, axis), lo(src, axis) - 1])
      }
      const axis = area.kind === 'cell' ? null : endAxis(x, src)
      // the cells move along the axis of the range's end they hold; moving across it does not adjust the range
      if (axis && (axis === 'r' ? dc : dr) === 0) return adjusted(sheet, area, axis, moveSpan(lo(x, axis), hi(x, axis), lo(src, axis), hi(src, axis), axis === 'r' ? dr : dc))
      return onDst ? overwritten(sheet, area, x, dst, onSrc) : undefined
    },
  }
}

/**
 * Names local to `from` that formulas moved or copied to `to` use, where `to` sees no name of that name
 * (local or global): Excel defines them on `to` too, `repoint` gives the reference.
 */
export function carryNames(names: NameInput[], formulas: string[], from: string, to: string, repoint: (ref: string) => string): NameInput[] {
  const fk = sheetKey(from)
  const own = (name: string) => names.find((n) => n.scope !== undefined && sheetKey(n.scope) === fk && caseKey(n.name) === caseKey(name))
  const added: NameInput[] = []
  for (const f of formulas) {
    let ast
    try {
      ast = parse(f)
    } catch {
      continue
    }
    walk(ast, (n) => {
      if (n.k !== 'name') return
      const local = own(n.name)
      if (local && !inputNameFor(names, to, n.name) && !added.some((a) => caseKey(a.name) === caseKey(n.name))) added.push({ name: local.name, ref: repoint(local.ref), scope: to })
    })
  }
  return added
}

/** Cut `src` on sheet `from` and paste it with its top-left cell at row `row`, column `col` of `to`. */
export function moveRange(book: BookInput, from: string, src: Rect, to: string, row: number, col: number): Applied {
  const [dr, dc] = [row - src.r1, col - src.c1]
  const dst = shiftRect(src, dr, dc)
  if (dst.r2 >= MAX_ROWS || dst.c2 >= MAX_COLS) return { ok: false, code: 'off_sheet', message: say(book.locale, { key: 'moveOffSheet', sheet: to, cell: formatAddress(row, col) }) }
  const paste: Paste = { from, src, to, area: dst, offsets: [[dr, dc]], move: true }
  const refusal = pasteRefusal(book, paste)
  if (refusal) return refusal
  const [fk, tk] = [sheetKey(from), sheetKey(to)]
  const [fi, ti] = [sheetIndex(book, from), sheetIndex(book, to)]
  const geometry: Geometry = [{ kind: 'range', from: fi, src, to: ti, dr, dc }]
  const t = moveTransform(from, src, to, dr, dc)
  const names = sheetNames(book)
  const moved: Record<string, CellInput> = {}
  const formulas: string[] = []
  const place = placer(geometry)
  const p: Place = { s: 0, r: 0, c: 0 }
  const sheets = book.sheets.map((s, si) => {
    const k = sheetKey(s.name)
    let cells: Record<string, CellInput> = {}
    let changed = si === fi || si === ti
    if (changed) {
      // the cells where the geometry takes them: the moved ones to `to`, the rest of the sheet but those landed on stay
      const g = gridOf(s)
      for (let i = 0; i < g.addr.length; i++) {
        const addr = g.addr[i]
        const cell = g.cell[i]
        p.s = si
        p.r = g.row[i]
        p.c = g.col[i]
        const moves = si === fi && hasCell(src, p.r, p.c)
        if (!place(p)) continue
        if (moves) {
          const at = formatAddress(p.r, p.c)
          moved[at] = relocated(rewriteCell(cell, s.name, to, t, names), addr, at)
          if (cell.formula !== undefined) formulas.push(cell.formula)
        } else cells[addr] = rewriteCell(cell, s.name, s.name, t, names)
      }
    } else {
      cells = rewriteCells(s.cells, s.name, s.name, t, names)
      changed = cells !== s.cells
    }
  // a table the move takes whole goes to `to` with its cells (`tables/moves.ts`), the others stay on their sheet
    const tables = rewriteTemplates(s.tables, s.name, (ti) => (k === fk && contains(src, parseRange(ti.range)!) ? to : s.name), t, names)
    return changed || tables !== s.tables ? { ...s, cells, ...(tables ? { tables } : {}) } : s
  })
  Object.assign(sheets[ti].cells, moved)
  // the looks of the cells without an entry move by rectangles, the place they land on cleared
  const [cut, onto] = [layerOf(book.sheets[fi]), layerOf(book.sheets[ti])]
  if (cut || onto) {
    const [lf, lt] = movedLooks(cut, src, onto, dr, dc, fi === ti)
    if (lf) sheets[fi] = withLayer(sheets[fi], sheets[fi].cells, lf)
    sheets[ti] = withLayer(sheets[ti], sheets[ti].cells, lt!)
  }
  const out: BookInput = { ...book, sheets }
  let nm = rewriteNames(book.names, t, names)
  if (nm && fk !== tk) {
    const carried = carryNames(nm, formulas, from, to, (ref) => ref)
    if (carried.length) nm = [...nm, ...carried]
  }
  if (nm) out.names = nm
  pasteValidations(book, out, paste, t, names)
  pasteLayout(book, out, paste)
  moveConditional(book, out, paste, t, names)
  return { ok: true, book: out, geometry }
}
